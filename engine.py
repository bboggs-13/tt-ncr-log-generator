"""TTFE Inspection Log — import engine.

Turns a PlanRadar CSV export into a standard TT Excel log. Supports multiple
inspection types, each with its own template and column layout but identical
formatting (logo, borders, 10pt font, Status dropdown + filter + live
conditional highlighting, audit sheet).

Add or change a type by editing the PROFILES dict below — no UI changes needed.
"""
from __future__ import annotations

import colorsys
import io
from copy import copy
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule

ASSETS = Path(__file__).parent / "assets"
LOGO_PATH = ASSETS / "tt_logo.png"

# --- Shared layout (same for every template) ---
SHEET = "Sheet1"
HEADER_ROW = 7
DATA_START_ROW = 8
TITLE_ROW = 4
DATE_ROW = 5
FONT_SIZE = 10.0

# Status fills / options / highlighting
OPEN_FILL = PatternFill(start_color="FFBDD7EE", end_color="FFBDD7EE", fill_type="solid")
CLOSED_FILL = PatternFill(start_color="FFC6E0B4", end_color="FFC6E0B4", fill_type="solid")
INPROGRESS_FILL = PatternFill(start_color="FFFFE699", end_color="FFFFE699", fill_type="solid")
STATUS_OPTIONS = ["Open", "In Progress", "Closed"]

# Time Sensitive NCR Log: age bands measured from the log's "Date Updated" (C5).
# (min_days, max_days or None, label, fill) — muted, incrementally stronger tones.
AGE_BANDS = [
    (30, 60, "30–60 days", PatternFill(start_color="FFFFF2CC", end_color="FFFFF2CC", fill_type="solid")),
    (60, 90, "60–90 days", PatternFill(start_color="FFF8CBAD", end_color="FFF8CBAD", fill_type="solid")),
    (90, None, "90+ days", PatternFill(start_color="FFF0A8A8", end_color="FFF0A8A8", fill_type="solid")),
]
DATE_FMT = "yyyy-mm-dd"
TEXT_FIELDS = ("Description", "Recommended Action", "Remarks")  # left + vertically centered
CLOSED_WORDS = ("closed", "close")


def _pastel(hex6: str, saturation: float = 0.65, lightness: float = 0.86) -> PatternFill:
    """Pastel version of a brand colour: keep its hue, set saturation and lightness."""
    r, g, b = (int(hex6[i:i + 2], 16) / 255 for i in (0, 2, 4))
    hue, _l, _s = colorsys.rgb_to_hls(r, g, b)
    t = "".join(f"{round(c * 255):02X}" for c in colorsys.hls_to_rgb(hue, lightness, saturation))
    return PatternFill(start_color="FF" + t, end_color="FF" + t, fill_type="solid")


# Project Closeout NCR Log: rows shaded by NCR type, pastel TT primary colours
# (Olive #8B9064, Terracotta #D4451D, Dark Blue #004050).
TYPE_FILLS = [
    ("TR-1", "Olive", _pastel("8B9064")),
    ("TR-8", "Terracotta", _pastel("D4451D")),
    ("QA/QC", "Dark Blue", _pastel("004050")),
]

# Borders
_THIN = Side(style="thin", color="FF000000")
_THICK = Side(style="thick", color="FF000000")
THIN_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
THICK_BORDER = Border(left=_THICK, right=_THICK, top=_THICK, bottom=_THICK)

LOGO_WIDTH_IN = 3.0
EMU_PER_IN = 914400

# Alternate CSV header names accepted for a given logical field (first match wins).
ALIASES = {
    "Report #": ["Report", "Report No", "Report No.", "Report Number", "Report#"],
    "Non-Conformance #": ["Item #", "Item", "Observation #", "Inspection #", "Issue #"],
    "Type": ["NCR Type", "Type of NCR", "Category"],
    "Responsible Party": ["Responsible", "Party Responsible", "Responsible party",
                          "Trade", "Contractor", "Assigned to"],
}


@dataclass(frozen=True)
class Profile:
    key: str
    label: str
    template: str                 # filename in assets/
    column_map: dict              # csv field -> template column letter
    required_columns: list        # logical field names that must resolve
    center_columns: list          # 1-based column indices to center
    status_col: int               # 1-based index of the Status column
    last_col: int                 # 1-based index of the last column
    age_highlight: bool = False   # Time Sensitive: shade open rows by age of Observation Date
    type_highlight: bool = False  # Project Closeout: shade rows by NCR Type
    status_colors: bool = True    # blue/yellow/green Status-cell colours (off when rows are shaded)
    filter_open: bool = False     # open the log pre-filtered to non-Closed items
    legend_cols: tuple = ("E", "F")  # legend: (swatch column, label column)
    ui_note: str = ""             # short explanation shown in the app


_AGE_NOTE = ("Open items are shaded across the whole row by days since the Observation Date, "
             "measured from the log's Date Updated: under 30 none · 30–59 yellow · 60–89 orange · "
             "90+ red. Closed items are not shaded. The log opens filtered to open items — use the "
             "Status filter to show closed items.")

# Special Inspections columns (Type and Priority removed)
_SPECIAL_MAP = {
    "Report #": "A", "Non-Conformance #": "B", "Observation Date": "C",
    "Description": "D", "Recommended Action": "E", "Status": "F",
    "Remarks": "G", "Date Closed": "H", "Link to ticket": "I",
}
_SPECIAL_REQ = ["Report #", "Non-Conformance #", "Observation Date", "Status"]
_SPECIAL_CENTER = [1, 2, 3, 6, 8, 9]

_REGULAR_MAP = {
    "Report #": "A", "Non-Conformance #": "B", "Observation Date": "C",
    "Responsible Party": "D", "Description": "E", "Recommended Action": "F",
    "Status": "G", "Remarks": "H", "Date Closed": "I",
}
_REGULAR_REQ = ["Observation Date", "Status"]
_REGULAR_CENTER = [1, 2, 3, 4, 7, 9]


PROFILES = {
    "special": Profile(
        key="special",
        label="Special Inspections",
        template="TTFE_Template.xlsx",
        column_map=_SPECIAL_MAP, required_columns=_SPECIAL_REQ,
        center_columns=_SPECIAL_CENTER, status_col=6, last_col=9,
    ),
    "special_time": Profile(
        key="special_time",
        label="Special Inspections — Time Sensitive NCR Log",
        template="TTFE_Template_TimeSensitive.xlsx",
        column_map=_SPECIAL_MAP, required_columns=_SPECIAL_REQ,
        center_columns=_SPECIAL_CENTER, status_col=6, last_col=9,
        age_highlight=True, status_colors=False, filter_open=True,
        legend_cols=("F", "G"),  # narrow Status column holds the swatches
        ui_note=_AGE_NOTE,
    ),
    "special_closeout": Profile(
        key="special_closeout",
        label="Special Inspections — Project Closeout NCR Log",
        template="TTFE_Template_Closeout.xlsx",
        column_map={
            "Report #": "A", "Non-Conformance #": "B", "Observation Date": "C",
            "Type": "D", "Description": "E", "Recommended Action": "F",
            "Status": "G", "Remarks": "H", "Date Closed": "I", "Link to ticket": "J",
        },
        required_columns=["Report #", "Non-Conformance #", "Observation Date", "Type", "Status"],
        center_columns=[1, 2, 3, 4, 7, 9, 10], status_col=7, last_col=10,
        type_highlight=True, status_colors=False,
        legend_cols=("D", "E"),
        ui_note=("Rows are shaded by NCR Type in pastel TT colours: TR-1 olive · TR-8 terracotta · "
                 "QA/QC dark blue. Other or blank types are not shaded."),
    ),
    "regular": Profile(
        key="regular",
        label="Regular Site Inspections",
        template="TTFE_Template_Regular.xlsx",
        column_map=_REGULAR_MAP, required_columns=_REGULAR_REQ,
        center_columns=_REGULAR_CENTER, status_col=7, last_col=9,
    ),
    "regular_time": Profile(
        key="regular_time",
        label="Regular Site Inspections — Time Sensitive Site Inspection Log",
        template="TTFE_Template_Regular_TimeSensitive.xlsx",
        column_map=_REGULAR_MAP, required_columns=_REGULAR_REQ,
        center_columns=_REGULAR_CENTER, status_col=7, last_col=9,
        age_highlight=True, status_colors=False, filter_open=True,
        legend_cols=("D", "E"),  # narrow Responsible Party column holds the swatches
        ui_note=_AGE_NOTE,
    ),
}
DEFAULT_PROFILE = "special"

# What the UI shows: inspection types, and (optionally) log variants within a type.
INSPECTION_TYPES = {"special": "Special Inspections", "regular": "Regular Site Inspections"}
LOG_VARIANTS = {
    "special": {"Standard NCR Log": "special", "Time Sensitive NCR Log": "special_time",
                "Project Closeout NCR Log": "special_closeout"},
    "regular": {"Standard Site Inspection Log": "regular",
                "Time Sensitive Site Inspection Log": "regular_time"},
}


def resolve_profile_key(inspection_type: str, log_variant: str | None = None) -> str:
    variants = LOG_VARIANTS.get(inspection_type)
    if variants and log_variant in variants:
        return variants[log_variant]
    return inspection_type


def get_profile(key: str) -> Profile:
    return PROFILES.get(key, PROFILES[DEFAULT_PROFILE])


def _resolve_field(field_name: str, df: pd.DataFrame):
    """Return the actual CSV column for a logical field, honoring aliases."""
    if field_name in df.columns:
        return field_name
    for alt in ALIASES.get(field_name, []):
        if alt in df.columns:
            return alt
    return None


def _coerce(val):
    s = str(val).replace(",", "").replace("$", "").strip()
    if s == "":
        return None
    try:
        f = float(s)
        return int(f) if f.is_integer() else f
    except ValueError:
        return val


def expected_columns(profile: Profile) -> list:
    return list(profile.column_map.keys())


def validate(df: pd.DataFrame, profile: Profile):
    msgs = []
    missing = [c for c in profile.required_columns if _resolve_field(c, df) is None]
    if missing:
        msgs.append(f"Missing required columns: {', '.join(missing)}")
    if df.empty:
        msgs.append("The CSV has no data rows.")
    mapped = set()
    for fld in profile.column_map:
        actual = _resolve_field(fld, df)
        if actual:
            mapped.add(actual)
    unknown = [c for c in df.columns if c not in mapped]
    if unknown:
        msgs.append(f"Note: unmapped columns will be ignored: {', '.join(unknown)}")
    ok = not missing and not df.empty
    return ok, msgs


def _apply_age_highlight(ws, profile, n, last_row, last_col_letter):
    """Shade whole data rows by days between Observation Date and Date Updated (C5).

    C5 and the Observation Date cells are written as real dates so Excel can do the
    math; editing C5 later re-ages the log automatically.
    """
    gen = date.today()
    ws["C5"] = gen
    ws["C5"].number_format = DATE_FMT

    obs_letter = profile.column_map["Observation Date"]
    for r in range(DATA_START_ROW, DATA_START_ROW + n):
        cell = ws[f"{obs_letter}{r}"]
        parsed = pd.to_datetime(cell.value, errors="coerce") if cell.value not in (None, "") else pd.NaT
        if not pd.isna(parsed):
            cell.value = parsed.date()
            cell.number_format = DATE_FMT

    if n > 0:
        rng = f"A{DATA_START_ROW}:{last_col_letter}{last_row}"
        age = f"($C$5-${obs_letter}{DATA_START_ROW})"
        # Only items that are still open get aged; Closed items are never shaded.
        st_cell = f"LOWER(TRIM(${get_column_letter(profile.status_col)}{DATA_START_ROW}))"
        not_closed = ",".join(f'{st_cell}<>"{w}"' for w in CLOSED_WORDS)
        for lo, hi, _label, fill in AGE_BANDS:
            cond = f"{age}>={lo}" if hi is None else f"{age}>={lo},{age}<{hi}"
            ws.conditional_formatting.add(rng, FormulaRule(
                formula=[f"AND(ISNUMBER(${obs_letter}{DATA_START_ROW}),{not_closed},{cond})"], fill=fill))

    rows = [("Less than 30 days (no highlight)", None)] + [
        (f"{lo} to {hi - 1} days" if hi else f"{lo}+ days", fill) for lo, hi, _l, fill in AGE_BANDS
    ]
    _write_legend(ws, profile,
                  "Legend: Days Since Observation Date, Open Items Only (as of Date Updated)", rows)


def _apply_type_highlight(ws, profile, n, last_row, last_col_letter):
    """Shade whole data rows by NCR Type (TR-1 / TR-8 / QA/QC)."""
    t = profile.column_map["Type"]
    if n > 0:
        rng = f"A{DATA_START_ROW}:{last_col_letter}{last_row}"
        for value, _name, fill in TYPE_FILLS:
            ws.conditional_formatting.add(rng, FormulaRule(
                formula=[f'EXACT(UPPER(TRIM(${t}{DATA_START_ROW})),"{value}")'], fill=fill, stopIfTrue=True))
    rows = [(v, fill) for v, _name, fill in TYPE_FILLS] + [("Other / blank type (no highlight)", None)]
    _write_legend(ws, profile, "Legend: NCR Type", rows)


def _apply_open_filter(ws, profile, n, last_row):
    """Pre-apply the Status filter so only non-Closed items show on open.

    Closed rows are hidden by the filter (not deleted) — choosing 'Select All' in the
    Status dropdown brings them back. Skipped if nothing is closed or everything is.
    """
    letter = get_column_letter(profile.status_col)
    shown, hidden = [], []
    for r in range(DATA_START_ROW, last_row + 1):
        v = ws[f"{letter}{r}"].value
        text = "" if v is None else str(v)
        if text.strip().lower() in CLOSED_WORDS:
            hidden.append(r)
        elif text not in shown:
            shown.append(text)
    if not hidden or len(hidden) == n:
        return 0
    ws.auto_filter.add_filter_column(profile.status_col - 1, [t for t in shown if t],
                                     blank="" in shown)
    for r in hidden:
        ws.row_dimensions[r].hidden = True
    return len(hidden)


def _write_legend(ws, profile, title, rows):
    """Colour key in rows 1-5: swatch in one column, meaning in the next.

    Column pair comes from Profile.legend_cols; fills come from the same constants
    as the highlighting rules, so the key always matches the rows.
    """
    sw_col, lab_col = profile.legend_cols
    head = ws[f"{sw_col}1"]
    head.value = title
    head.font = Font(name="Arial", size=10, bold=True)
    head.alignment = Alignment(horizontal="left", vertical="center")
    head.border = THIN_BORDER
    ws.merge_cells(f"{sw_col}1:{lab_col}1")

    for r, (text, fill) in enumerate(rows, start=2):
        swatch, label = ws[f"{sw_col}{r}"], ws[f"{lab_col}{r}"]
        swatch.border = THIN_BORDER
        if fill is not None:
            swatch.fill = fill
        label.value = text
        label.font = Font(name="Arial", size=10)
        label.alignment = Alignment(horizontal="left", vertical="center")
        label.border = THIN_BORDER


def build_workbook(df: pd.DataFrame, profile: Profile, project_name: str = "134 Jane Street") -> bytes:
    wb = load_workbook(ASSETS / profile.template)
    ws = wb[SHEET]
    ws._images = []  # we place exactly one logo

    first_col, last_col = 1, profile.last_col
    status_col = profile.status_col

    # Header cells
    ws["A4"] = project_name
    ws["C5"] = date.today().strftime("%Y-%m-%d")

    # Data rows (resolve each field through aliases)
    n = len(df)
    resolved = {fld: _resolve_field(fld, df) for fld in profile.column_map}
    for i, (_, row) in enumerate(df.iterrows()):
        r = DATA_START_ROW + i
        for fld, col_letter in profile.column_map.items():
            src = resolved[fld]
            if src is not None:
                ws[f"{col_letter}{r}"] = _coerce(row[src])

    # Center fixed columns
    for i in range(n):
        r = DATA_START_ROW + i
        for c in profile.center_columns:
            cell = ws.cell(r, c)
            al = copy(cell.alignment)
            cell.alignment = Alignment(horizontal="center", vertical="center",
                                       wrap_text=al.wrap_text, text_rotation=al.text_rotation,
                                       indent=al.indent)

    # Text columns: left aligned, vertically centered, wrapped
    text_cols = [profile.column_map[f] for f in TEXT_FIELDS if f in profile.column_map]
    for i in range(n):
        r = DATA_START_ROW + i
        for col_letter in text_cols:
            cell = ws[f"{col_letter}{r}"]
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # Borders: thin on all data cells; thick on header row only
    for r in range(DATA_START_ROW, DATA_START_ROW + n):
        for c in range(first_col, last_col + 1):
            ws.cell(r, c).border = THIN_BORDER
    for c in range(first_col, last_col + 1):
        ws.cell(HEADER_ROW, c).border = THICK_BORDER

    # Font: 10pt everywhere (title, headers, data), preserving bold/name/color
    for r in range(TITLE_ROW, DATA_START_ROW + n):
        for c in range(first_col, last_col + 1):
            cell = ws.cell(r, c)
            f = cell.font
            cell.font = Font(name=f.name, size=FONT_SIZE, bold=f.bold,
                             italic=f.italic, color=f.color, underline=f.underline)

    # Filter on the header row
    last_row = DATA_START_ROW + n - 1
    last_col_letter = get_column_letter(last_col)
    ws.auto_filter.ref = f"A{HEADER_ROW}:{last_col_letter}{last_row}"

    # Dropdown on the Status column
    status_letter = get_column_letter(status_col)
    dv = DataValidation(type="list", formula1='"' + ",".join(STATUS_OPTIONS) + '"',
                        allow_blank=True, showDropDown=False)
    dv.error = "Choose Open, In Progress, or Closed."
    dv.errorTitle = "Invalid status"
    dv.prompt = "Select a status"
    ws.add_data_validation(dv)
    dv.add(f"{status_letter}{DATA_START_ROW}:{status_letter}{last_row}")

    # Live Status-cell highlighting (recolors on dropdown change). Turned off for logs
    # whose rows are shaded by another rule, so that colour runs across the whole row.
    if profile.status_colors:
        status_range = f"{status_letter}{DATA_START_ROW}:{status_letter}{last_row}"
        top = f"{status_letter}{DATA_START_ROW}"
        ws.conditional_formatting.add(status_range,
            FormulaRule(formula=[f'EXACT(LOWER(TRIM({top})),"open")'], fill=OPEN_FILL, stopIfTrue=True))
        ws.conditional_formatting.add(status_range,
            FormulaRule(formula=[f'EXACT(LOWER(TRIM({top})),"in progress")'], fill=INPROGRESS_FILL, stopIfTrue=True))
        ws.conditional_formatting.add(status_range,
            FormulaRule(formula=[f'EXACT(LOWER(TRIM({top})),"closed")'], fill=CLOSED_FILL, stopIfTrue=True))

    if profile.age_highlight:
        _apply_age_highlight(ws, profile, n, last_row, last_col_letter)
    if profile.type_highlight:
        _apply_type_highlight(ws, profile, n, last_row, last_col_letter)
    hidden_closed = _apply_open_filter(ws, profile, n, last_row) if profile.filter_open and n else 0

    # Logo, exactly 3in wide, top-left corner
    from PIL import Image as PILImage
    w_px, h_px = PILImage.open(LOGO_PATH).size
    width_emu = int(EMU_PER_IN * LOGO_WIDTH_IN)
    height_emu = int(width_emu * h_px / w_px)
    img = XLImage(str(LOGO_PATH))
    img.anchor = OneCellAnchor(_from=AnchorMarker(col=0, colOff=0, row=0, rowOff=0),
                               ext=XDRPositiveSize2D(width_emu, height_emu))
    ws.add_image(img)

    # Audit sheet
    audit = wb.create_sheet("Import Audit")
    for idx, (k, v) in enumerate([
        ("Generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Inspection type", profile.label),
        ("Aging reference date", ws["C5"].value.strftime("%Y-%m-%d") if profile.age_highlight else "n/a"),
        ("Template", profile.template),
        ("Default filter", f"Open items only ({hidden_closed} closed hidden)" if hidden_closed else "None"),
        ("Rows imported", n),
        ("Project", project_name),
    ], start=1):
        audit[f"A{idx}"] = k
        audit[f"B{idx}"] = v

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def process_csv_bytes(csv_bytes: bytes, inspection_type: str = DEFAULT_PROFILE,
                      project_name: str = "134 Jane Street"):
    """Full pipeline. Returns (xlsx_bytes_or_None, df, messages)."""
    profile = get_profile(inspection_type)
    df = pd.read_csv(io.BytesIO(csv_bytes), dtype=str).fillna("")
    ok, msgs = validate(df, profile)
    if not ok:
        return None, df, msgs
    xlsx = build_workbook(df, profile, project_name=project_name)
    return xlsx, df, msgs
