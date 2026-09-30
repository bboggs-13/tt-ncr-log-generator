"""TTFE NCR Log Generator — Streamlit web app.

Upload a PlanRadar NCR / Special-Inspections CSV, download the formatted
Thornton Tomasetti Excel log. Template, logo, mapping, and formatting are
built in — the user only provides the CSV.

Styling mirrors the Thornton Tomasetti website: warm cream canvas, bold
headings, olive (#8B9064) + terra cotta (#D4451D) accents, cool-gray body
text (#76767B). Theme base lives in .streamlit/config.toml.

Run locally:   streamlit run app.py
"""
from pathlib import Path

import streamlit as st
from PIL import Image

import base64

import engine

_ICON = Image.open(Path(engine.ASSETS) / "tt_icon.png")

def _icon_b64(name):
    return base64.b64encode((Path(engine.ASSETS) / name).read_bytes()).decode()

_PAPERCLIP_B64 = _icon_b64("paperclip.png")
_BUILDING_B64 = _icon_b64("building.png")

def field_label(icon_b64, text):
    st.markdown(
        f'<div class="tt-field-label">'
        f'<img src="data:image/png;base64,{icon_b64}"/> {text}</div>',
        unsafe_allow_html=True,
    )

st.set_page_config(
    page_title="TT NCR Log Generator",
    page_icon=_ICON,
    layout="centered",
)

# --- TT-website-inspired styling ---
st.markdown(
    """
    <style>
      :root {
        --tt-olive: #8B9064;
        --tt-terra: #D4451D;
        --tt-gray:  #76767B;
        --tt-ink:   #1A1A17;
        --tt-cream: #FAF9F5;
      }
      .block-container { padding-top: 3.5rem; max-width: 860px; }

      /* Small uppercase category label, like the TT site's NEWS / INSIGHT tags */
      .tt-eyebrow {
        color: var(--tt-terra);
        font-size: 0.75rem;
        font-weight: 700;
        letter-spacing: 0.18em;
        text-transform: uppercase;
        margin: 0.75rem 0 0.25rem;
      }
      /* Big bold near-black headline, TT-style */
      h1 {
        color: var(--tt-ink) !important;
        font-weight: 800 !important;
        letter-spacing: -0.02em;
        line-height: 1.05;
      }
      h2, h3 {
        color: var(--tt-ink) !important;
        font-weight: 700 !important;
        letter-spacing: -0.01em;
      }
      /* Section labels get an olive tint to echo the wordmark */
      h3 { color: var(--tt-olive) !important; text-transform: none; }

      /* Terra-cotta rule under the header, echoing the site's accent bars */
      .tt-rule { height: 4px; width: 64px; background: var(--tt-terra);
                 border: none; margin: 1rem 0 1.5rem; border-radius: 2px; }

      /* Metric values in terra cotta, labels in olive */
      [data-testid="stMetricValue"] { color: var(--tt-terra) !important; font-weight: 800; }
      [data-testid="stMetricLabel"] { color: var(--tt-olive) !important;
                 text-transform: uppercase; letter-spacing: 0.08em; font-size: 0.72rem; }

      /* Buttons: use primaryColor (terra cotta); darken slightly on hover */
      .stDownloadButton button, .stButton button { font-weight: 700; letter-spacing: 0.01em; }
      .stDownloadButton button:hover, .stButton button:hover { filter: brightness(0.93); }

      /* Icon + text labels for the input fields */
      .tt-field-label {
        display: flex; align-items: center; gap: 8px;
        font-size: 0.9rem; font-weight: 600; color: var(--tt-ink);
        margin: 0.5rem 0 0.35rem;
      }
      .tt-field-label img { height: 20px; width: 20px; object-fit: contain; }

      /* Keep the logo fully visible: square edges, no crop */
      [data-testid="stImage"] img, .stImage img {
        border-radius: 0 !important;
        object-fit: contain !important;
        clip-path: none !important;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

# --- Header ---
st.image(str(engine.LOGO_PATH), width=280)
st.markdown('<div class="tt-eyebrow">Façade Engineering · Quality Control</div>', unsafe_allow_html=True)
st.title("NCR Log Generator")
st.markdown('<hr class="tt-rule">', unsafe_allow_html=True)
st.caption(
    "Upload a PlanRadar Non-Conformance / Special-Inspections CSV export. "
    "The app fills the standard TT template — formatting, Status highlighting, "
    "column centering, and logo included — and returns a ready-to-send Excel file."
)

with st.expander("Expected CSV columns"):
    st.write(", ".join(engine.COLUMN_MAP.keys()))
    st.write(f"**Required:** {', '.join(engine.REQUIRED_COLUMNS)}")

field_label(_BUILDING_B64, "Project name (written to cell A4)")
project = st.text_input(
    "Project name (written to cell A4)",
    value="134 Jane Street",
    label_visibility="collapsed",
)

field_label(_PAPERCLIP_B64, "Upload CSV")
uploaded = st.file_uploader(
    "Upload CSV", type=["csv"], label_visibility="collapsed"
)

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
