"""TTFE Inspection Log — import engine.

Turns a PlanRadar CSV export into a standard TT Excel log. Supports multiple
inspection types, each with its own template and column layout but identical
formatting (logo, borders, 10pt font, Status dropdown + filter + live
conditional highlighting, audit sheet).

Add or change a type by editing the PROFILES dict below — no UI changes needed.
"""
from __future__ import annotations

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


PROFILES = {
    "special": Profile(
        key="special",
        label="Special Inspections",
        template="TTFE_Template.xlsx",
        column_map={
            "Report #": "A", "Non-Conformance #": "B", "Observation Date": "C",
            "Type": "D", "Priority": "E", "Description": "F", "Recommended Action": "G",
            "Status": "H", "Remarks": "I", "Date Closed": "J", "Link to ticket": "K",
        },
        required_columns=["Report #", "Non-Conformance #", "Observation Date", "Type", "Status"],
        center_columns=[1, 2, 3, 4, 5, 8, 10, 11],
        status_col=8,
        last_col=11,
    ),
    "regular": Profile(
        key="regular",
        label="Regular Site Inspections",
        template="TTFE_Template_Regular.xlsx",
        column_map={
            "Report #": "A", "Non-Conformance #": "B", "Observation Date": "C",
            "Responsible Party": "D", "Description": "E", "Recommended Action": "F",
            "Status": "G", "Remarks": "H", "Date Closed": "I",
        },
        required_columns=["Observation Date", "Status"],
        center_columns=[1, 2, 3, 4, 7, 9],
        status_col=7,
        last_col=9,
    ),
}
DEFAULT_PROFILE = "special"


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

    # Live highlighting via conditional formatting (recolors on dropdown change)
    status_range = f"{status_letter}{DATA_START_ROW}:{status_letter}{last_row}"
    top = f"{status_letter}{DATA_START_ROW}"
    ws.conditional_formatting.add(status_range,
        FormulaRule(formula=[f'EXACT(LOWER(TRIM({top})),"open")'], fill=OPEN_FILL, stopIfTrue=True))
    ws.conditional_formatting.add(status_range,
        FormulaRule(formula=[f'EXACT(LOWER(TRIM({top})),"in progress")'], fill=INPROGRESS_FILL, stopIfTrue=True))
    ws.conditional_formatting.add(status_range,
        FormulaRule(formula=[f'EXACT(LOWER(TRIM({top})),"closed")'], fill=CLOSED_FILL, stopIfTrue=True))

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
        ("Template", profile.template),
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
