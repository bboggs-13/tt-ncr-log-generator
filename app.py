"""TTFE NCR Log Generator — Streamlit web app.

Upload a PlanRadar NCR / Special-Inspections CSV, download the formatted
Thornton Tomasetti Excel log. Template, logo, mapping, and formatting are
built in — the user only provides the CSV.

Theme: Anthropic-inspired warm light look with TT brand colors
(olive #8B9064, terra cotta #D4451D, cool-gray text #76767B) in
.streamlit/config.toml. This file adds a few brand accents on top.

Run locally:   streamlit run app.py
"""
import streamlit as st

import engine

st.set_page_config(page_title="TT NCR Log Generator", page_icon="🧱", layout="centered")

# --- Brand accents (headings olive, title terra cotta, warm spacing) ---
st.markdown(
    """
    <style>
      :root {
        --tt-olive: #8B9064;
        --tt-terra: #D4451D;
        --tt-gray:  #76767B;
      }
      .block-container { padding-top: 3rem; max-width: 820px; }
      h1 { color: var(--tt-terra) !important; font-weight: 700; letter-spacing: -0.01em; }
      h2, h3 { color: var(--tt-olive) !important; font-weight: 600; }
      /* Thin olive rule under the header block */
      .tt-divider { height: 3px; background: var(--tt-olive); border: none;
                    opacity: 0.35; margin: 0.25rem 0 1.5rem; border-radius: 2px; }
      /* Metric values in terra cotta */
      [data-testid="stMetricValue"] { color: var(--tt-terra) !important; }
      /* Primary buttons already use terra cotta via primaryColor;
         make hover a touch darker */
      .stDownloadButton button:hover, .stButton button:hover { filter: brightness(0.93); }
    </style>
    """,
    unsafe_allow_html=True,
)

# --- Header ---
st.image(str(engine.LOGO_PATH), width=260)
st.title("NCR Log Generator")
st.markdown('<hr class="tt-divider">', unsafe_allow_html=True)
st.caption(
    "Upload a PlanRadar Non-Conformance / Special-Inspections CSV export. "
    "The app fills the standard TT template — formatting, Status highlighting, "
    "column centering, and logo included — and gives you a ready-to-send Excel file."
)

with st.expander("Expected CSV columns"):
    st.write(", ".join(engine.COLUMN_MAP.keys()))
    st.write(f"**Required:** {', '.join(engine.REQUIRED_COLUMNS)}")

project = st.text_input("Project name (written to cell A4)", value="134 Jane Street")

uploaded = st.file_uploader("Upload CSV", type=["csv"])

if uploaded is not None:
    csv_bytes = uploaded.getvalue()
    try:
        xlsx, df, msgs = engine.process_csv_bytes(csv_bytes, project_name=project.strip() or "134 Jane Street")
    except Exception as e:  # noqa: BLE001
        st.error(f"Could not read the CSV: {e}")
        st.stop()

    ok = xlsx is not None
    blocking = [m for m in msgs if not m.startswith("Note:")]
    notes = [m for m in msgs if m.startswith("Note:")]

    if not ok:
        for m in blocking:
            st.error(m)
        st.stop()

    for m in notes:
        st.info(m)

    st.success(f"Validated {len(df)} rows.")

    st.subheader("Preview")
    st.dataframe(df.head(15), use_container_width=True)

    # Status summary
    counts = df["Status"].str.strip().str.lower().value_counts().to_dict() if "Status" in df else {}
    open_n = counts.get("open", 0)
    inprog_n = counts.get("in progress", 0) + counts.get("in-progress", 0) + counts.get("wip", 0)
    closed_n = counts.get("closed", 0) + counts.get("close", 0)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total rows", len(df))
    c2.metric("Open", open_n)
    c3.metric("In Progress", inprog_n)
    c4.metric("Closed", closed_n)

    out_name = f"{(project.strip() or 'NCR').replace(' ', '_')}_NCR_Log.xlsx"
    st.download_button(
        "⬇ Download completed Excel workbook",
        data=xlsx,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
else:
    st.info("Awaiting a CSV upload to begin.")
