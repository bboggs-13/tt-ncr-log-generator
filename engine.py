"""TTFE NCR Log — import engine.

Turns a PlanRadar NCR/Special-Inspections CSV into the standard TT Excel log:
copies the master template, writes rows, sets header cells, highlights Status,
centers the fixed columns, and drops in a single 3-inch TT logo.

The template format is fixed, so the mapping lives here as constants.
"""
from __future__ import annotations

import io
from copy import copy
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

ASSETS = Path(__file__).parent / "assets"
TEMPLATE_PATH = ASSETS / "TTFE_Template.xlsx"
LOGO_PATH = ASSETS / "tt_logo.png"

SHEET = "Sheet1"
HEADER_ROW = 7
DATA_START_ROW = 8

# CSV column -> template column letter
COLUMN_MAP = {
    "Report #": "A",
    "Non-Conformance #": "B",
    "Observation Date": "C",
    "Type": "D",
    "Priority": "E",
    "Description": "F",
    "Recommended Action": "G",
    "Status": "H",
    "Remarks": "I",
    "Date Closed": "J",
    "Link to ticket": "K",
}
REQUIRED_COLUMNS = ["Report #", "Non-Conformance #", "Observation Date", "Type", "Status"]
CENTER_COLUMNS = [1, 2, 3, 4, 5, 8, 10, 11]  # A,B,C,D,E,H,J,K

STATUS_COL = 8  # H
OPEN_FILL = PatternFill(start_color="FFBDD7EE", end_color="FFBDD7EE", fill_type="solid")
CLOSED_FILL = PatternFill(start_color="FFC6E0B4", end_color="FFC6E0B4", fill_type="solid")
INPROGRESS_FILL = PatternFill(start_color="FFFFE699", end_color="FFFFE699", fill_type="solid")
STATUS_OPTIONS = ["Open", "In Progress", "Closed"]

# Borders: thin on every data cell, thick around the title and header row.
_THIN = Side(style="thin", color="FF000000")
_THICK = Side(style="thick", color="FF000000")
THIN_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
THICK_BORDER = Border(left=_THICK, right=_THICK, top=_THICK, bottom=_THICK)
FIRST_COL, LAST_COL = 1, 11  # A..K
FONT_SIZE = 10.0
TITLE_ROW = 4
DATE_ROW = 5

LOGO_WIDTH_IN = 3.0
EMU_PER_IN = 914400


def _coerce(val):
    """Numeric-looking strings -> numbers so Excel treats them as values."""
    s = str(val).replace(",", "").replace("$", "").strip()
    if s == "":
        return None
    try:
        f = float(s)
        return int(f) if f.is_integer() else f
    except ValueError:
        return val


def validate(df: pd.DataFrame):
    """Return (ok, messages). messages is a list of human-readable strings."""
    msgs = []
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        msgs.append(f"Missing required columns: {', '.join(missing)}")
    if df.empty:
        msgs.append("The CSV has no data rows.")
    unknown = [c for c in df.columns if c not in COLUMN_MAP]
    if unknown:
        msgs.append(f"Note: unmapped columns will be ignored: {', '.join(unknown)}")
    ok = not missing and not df.empty
    return ok, msgs


def build_workbook(df: pd.DataFrame, project_name: str = "134 Jane Street") -> bytes:
    """Populate the template from df and return xlsx bytes."""
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb[SHEET]

    # Clear any images the template ships with; we place exactly one logo.
    ws._images = []

    # Header cells
    ws["A4"] = project_name
    ws["C5"] = date.today().strftime("%Y-%m-%d")

    # Data rows
    n = len(df)
    for i, (_, row) in enumerate(df.iterrows()):
        r = DATA_START_ROW + i
        for field, col_letter in COLUMN_MAP.items():
            if field in df.columns:
                ws[f"{col_letter}{r}"] = _coerce(row[field])

    # Status highlight is applied via conditional formatting below so it
    # repaints live when a user changes the dropdown value in Excel.

    # Center fixed columns
    for i in range(n):
        r = DATA_START_ROW + i
        for c in CENTER_COLUMNS:
            cell = ws.cell(r, c)
            al = copy(cell.alignment)
            cell.alignment = Alignment(
                horizontal="center", vertical="center",
                wrap_text=al.wrap_text, text_rotation=al.text_rotation, indent=al.indent,
            )

    # Borders: thin on all data cells; thick on header row 7 only
    for r in range(DATA_START_ROW, DATA_START_ROW + n):
        for c in range(FIRST_COL, LAST_COL + 1):
            ws.cell(r, c).border = THIN_BORDER
    for c in range(FIRST_COL, LAST_COL + 1):
        ws.cell(HEADER_ROW, c).border = THICK_BORDER

    # Font: 10pt everywhere (title, headers, data), preserving bold/name/color
    for r in range(TITLE_ROW, DATA_START_ROW + n):
        for c in range(FIRST_COL, LAST_COL + 1):
            cell = ws.cell(r, c)
            f = cell.font
            cell.font = Font(name=f.name, size=FONT_SIZE, bold=f.bold,
                             italic=f.italic, color=f.color, underline=f.underline)

    # Filter on the header row so users can filter/sort by Status (and any column)
    last_row = DATA_START_ROW + n - 1
    last_col_letter = get_column_letter(LAST_COL)
    ws.auto_filter.ref = f"A{HEADER_ROW}:{last_col_letter}{last_row}"

    # Dropdown on the Status column: Open / In Progress / Closed
    status_letter = get_column_letter(STATUS_COL)
    dv = DataValidation(
        type="list",
        formula1='"' + ",".join(STATUS_OPTIONS) + '"',
        allow_blank=True,
        showDropDown=False,
    )
    dv.error = "Choose Open, In Progress, or Closed."
    dv.errorTitle = "Invalid status"
    dv.prompt = "Select a status"
    ws.add_data_validation(dv)
    dv.add(f"{status_letter}{DATA_START_ROW}:{status_letter}{last_row}")

    # Live highlighting: conditional formatting recolors the Status cell
    # instantly when the dropdown value changes in Excel (case-insensitive).
    status_range = f"{status_letter}{DATA_START_ROW}:{status_letter}{last_row}"
    top = f"{status_letter}{DATA_START_ROW}"
    ws.conditional_formatting.add(
        status_range,
        FormulaRule(formula=[f'EXACT(LOWER(TRIM({top})),"open")'], fill=OPEN_FILL, stopIfTrue=True),
    )
    ws.conditional_formatting.add(
        status_range,
        FormulaRule(formula=[f'EXACT(LOWER(TRIM({top})),"in progress")'], fill=INPROGRESS_FILL, stopIfTrue=True),
    )
    ws.conditional_formatting.add(
        status_range,
        FormulaRule(formula=[f'EXACT(LOWER(TRIM({top})),"closed")'], fill=CLOSED_FILL, stopIfTrue=True),
    )

    # Logo, exactly 3in wide, top-left corner
    from PIL import Image as PILImage
    w_px, h_px = PILImage.open(LOGO_PATH).size
    width_emu = int(EMU_PER_IN * LOGO_WIDTH_IN)
    height_emu = int(width_emu * h_px / w_px)
    img = XLImage(str(LOGO_PATH))
    img.anchor = OneCellAnchor(
        _from=AnchorMarker(col=0, colOff=0, row=0, rowOff=0),
        ext=XDRPositiveSize2D(width_emu, height_emu),
    )
    ws.add_image(img)

    # Audit sheet
    audit = wb.create_sheet("Import Audit")
    for idx, (k, v) in enumerate([
        ("Generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Template", TEMPLATE_PATH.name),
        ("Rows imported", n),
        ("Project", project_name),
    ], start=1):
        audit[f"A{idx}"] = k
        audit[f"B{idx}"] = v

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def process_csv_bytes(csv_bytes: bytes, project_name: str = "134 Jane Street"):
    """Full pipeline from raw CSV bytes. Returns (xlsx_bytes, df, messages)."""
    df = pd.read_csv(io.BytesIO(csv_bytes), dtype=str).fillna("")
    ok, msgs = validate(df)
    if not ok:
        return None, df, msgs
    xlsx = build_workbook(df, project_name=project_name)
    return xlsx, df, msgs
