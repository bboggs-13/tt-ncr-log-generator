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
import html

import engine

_ICON = Image.open(Path(engine.ASSETS) / "tt_icon.png")

def _icon_b64(name):
    return base64.b64encode((Path(engine.ASSETS) / name).read_bytes()).decode()

_PAPERCLIP_B64 = _icon_b64("paperclip.png")
_BUILDING_B64 = _icon_b64("building.png")
_MAGNIFIER_B64 = _icon_b64("clipboard.png")
_LINK_B64 = _icon_b64("link.png")

def field_label(icon_b64, text, tip=None, container=None):
    """Icon + text label for an input. With `tip`, an (i) icon is added and the
    explanation shows only while the mouse is over the label (or it has focus)."""
    tip_html = ""
    if tip:
        tip_html = ('<span class="tt-tip-icon" aria-hidden="true">i</span>'
                    f'<span class="tt-tip-text" role="tooltip">{html.escape(tip)}</span>')
    cls, attrs = ("tt-field-label tt-has-tip", ' tabindex="0"') if tip else ("tt-field-label", "")
    (container or st).markdown(
        f'<div class="{cls}"{attrs}>'
        f'<img src="data:image/png;base64,{icon_b64}"/> {text}{tip_html}</div>',
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

      /* Hover-only explanations on field labels: frosted cream card, olive accent */
      .tt-has-tip { position: relative; cursor: help; width: fit-content; outline: none; }
      .tt-tip-icon {
        display: inline-flex; align-items: center; justify-content: center;
        width: 16px; height: 16px; border-radius: 50%;
        background: rgba(139, 144, 100, 0.14); color: var(--tt-olive);
        font-size: 10px; font-weight: 700; font-style: italic; font-family: Georgia, serif;
        transition: background 0.15s ease, color 0.15s ease;
      }
      .tt-has-tip:hover .tt-tip-icon, .tt-has-tip:focus .tt-tip-icon {
        background: var(--tt-olive); color: #FFFFFF;
      }
      .tt-tip-text {
        visibility: hidden; opacity: 0; transform: translateY(-4px);
        transition: opacity 0.16s ease, transform 0.16s ease, visibility 0.16s;
        position: absolute; left: -6px; top: calc(100% + 10px); z-index: 1000;
        width: max-content; max-width: min(420px, 85vw);
        background: rgba(250, 249, 245, 0.88);           /* TT cream, slightly see-through */
        -webkit-backdrop-filter: blur(10px) saturate(1.2);
        backdrop-filter: blur(10px) saturate(1.2);
        border: 1px solid rgba(139, 144, 100, 0.28);
        border-left: 3px solid var(--tt-olive);
        border-radius: 10px;
        box-shadow: 0 8px 24px rgba(26, 26, 23, 0.10), 0 1px 3px rgba(26, 26, 23, 0.06);
        color: var(--tt-ink);
        font-size: 0.8rem; font-weight: 400; line-height: 1.55; letter-spacing: 0.005em;
        padding: 10px 14px 10px 12px;
        pointer-events: none;
      }
      .tt-tip-text::before {                               /* small pointer toward the label */
        content: ""; position: absolute; top: -6px; left: 18px;
        width: 10px; height: 10px; transform: rotate(45deg);
        background: rgba(250, 249, 245, 0.96);
        border-top: 1px solid rgba(139, 144, 100, 0.28);
        border-left: 1px solid rgba(139, 144, 100, 0.28);
        border-top-left-radius: 2px;
      }
      .tt-has-tip:hover .tt-tip-text, .tt-has-tip:focus .tt-tip-text {
        visibility: visible; opacity: 1; transform: translateY(0);
      }
      /* Let the tooltip float over the widget below instead of being clipped */
      [data-testid="stElementContainer"]:has(.tt-has-tip),
      .element-container:has(.tt-has-tip) { position: relative; z-index: 50; overflow: visible !important; }
      [data-testid="stMarkdownContainer"]:has(.tt-has-tip) { overflow: visible !important; }

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

# Inspection type selector (drives which template/columns are used)
field_label(_MAGNIFIER_B64, "Inspection type")
_type_keys = list(engine.INSPECTION_TYPES.keys())
_type_labels = dict(engine.INSPECTION_TYPES)
try:
    insp_type = st.segmented_control(
        "Inspection type", options=_type_keys,
        format_func=lambda k: _type_labels[k],
        default=engine.DEFAULT_PROFILE, label_visibility="collapsed",
    ) or engine.DEFAULT_PROFILE
except AttributeError:
    insp_type = st.radio(
        "Inspection type", options=_type_keys,
        format_func=lambda k: _type_labels[k],
        horizontal=True, label_visibility="collapsed",
    )

# PlanRadar access: Yes adds the "Link to ticket" column, No leaves it out
field_label(_LINK_B64, "Does the contractor have PlanRadar access?",
            tip="Yes includes the PlanRadar Link to ticket column in the log. "
                "No leaves it out, since the contractor can't open the links.")
_pr_options = ["Yes", "No"]
try:
    _pr = st.segmented_control(
        "PlanRadar access", options=_pr_options, default="Yes",
        label_visibility="collapsed", key="planradar_access",
    ) or "Yes"
except AttributeError:
    _pr = st.radio("PlanRadar access", options=_pr_options, horizontal=True,
                   label_visibility="collapsed", key="planradar_access")
planradar_access = _pr == "Yes"

# Log type buttons (same style as Inspection type) for each inspection type
log_variant = None
if insp_type in engine.LOG_VARIANTS:
    _log_label = st.empty()   # filled below, once the chosen log type is known
    _variants = list(engine.LOG_VARIANTS[insp_type].keys())
    try:
        log_variant = st.segmented_control(
            "Log type", options=_variants, default=_variants[0],
            label_visibility="collapsed", key=f"log_type_{insp_type}",
        ) or _variants[0]          # clicking the active button again keeps a selection
    except AttributeError:         # older Streamlit without segmented_control
        log_variant = st.radio(
            "Log type", options=_variants, horizontal=True,
            label_visibility="collapsed", key=f"log_type_{insp_type}",
        )
    _note = engine.get_profile(engine.resolve_profile_key(insp_type, log_variant)).ui_note
    field_label(_MAGNIFIER_B64, "Log type", tip=_note or None, container=_log_label)
profile_key = engine.resolve_profile_key(insp_type, log_variant)

_profile = engine.with_planradar_link(engine.get_profile(profile_key), planradar_access)
with st.expander(f"Expected CSV columns — {_profile.label}"):
    st.write(", ".join(engine.expected_columns(_profile)))
    st.write(f"**Required:** {', '.join(_profile.required_columns)}")

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
        xlsx, df, msgs = engine.process_csv_bytes(
            csv_bytes, inspection_type=profile_key,
            project_name=project.strip() or "134 Jane Street",
            planradar_access=planradar_access,
        )
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

    _suffix = {
        "special": "NCR_Log",
        "special_time": "Time_Sensitive_NCR_Log",
        "special_closeout": "Project_Closeout_NCR_Log",
        "regular": "Site_Inspection_Log",
        "regular_time": "Time_Sensitive_Site_Inspection_Log",
    }.get(profile_key, "Log")
    out_name = f"{(project.strip() or 'TT').replace(' ', '_')}_{_suffix}.xlsx"
    st.download_button(
        "⬇ Download completed Excel workbook",
        data=xlsx,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    # Time Sensitive logs: one-page PDF summary of open items by age band
    if _profile.age_highlight:
        import summary_pdf
        st.subheader("Summary PDF")
        _pdf = summary_pdf.build_summary_pdf(df, _profile, project.strip() or "134 Jane Street")
        st.download_button(
            "📄 Generate summary PDF",
            data=_pdf,
            file_name=out_name.replace(".xlsx", "_Summary.pdf"),
            mime="application/pdf",
            help="One-page summary of open items in each age category, with the project name "
                 "and today's date as the date of distribution.",
        )
else:
    st.info("Awaiting a CSV upload to begin.")
