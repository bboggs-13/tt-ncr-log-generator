"""Landscape PDF of an Excel log, showing exactly what the downloaded workbook shows.

Reads the workbook built by engine.build_workbook, so the PDF always matches it:
  • only the rows visible when the log opens (open items; closed rows are hidden),
  • the same columns in the same order, including Link to ticket when present,
  • the same colours: Status-cell fills, Time Sensitive age shading, Closeout type shading,
  • the same title block: TT logo, project name, Date Updated and the legend.

Page: US Letter landscape with 0.5 in margins on every side. All columns fit the page
width; rows continue onto further pages with the column headers repeated. Text is Arial
(or Liberation Sans, its metric-identical stand-in, when Arial isn't installed).
"""
from __future__ import annotations

import io
from datetime import date, datetime
from pathlib import Path
from xml.sax.saxutils import escape

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.platypus import BaseDocTemplate, Frame, Image, PageTemplate, Paragraph, Spacer, Table, TableStyle

import engine

PAGE = landscape(letter)                  # 792 x 612 pt
MARGIN = 36                               # 0.5 in on all four sides
CONTENT_W = PAGE[0] - 2 * MARGIN          # 720 pt
FOOTER_H = 14                             # footer sits inside the bottom margin box
BODY_SIZE, LEADING = 8.0, 9.6
PAD_X, PAD_Y = 3.0, 2.5
THIN, THICK = 0.5, 1.5                    # Excel thin / thick borders
MUTED = colors.HexColor("#6E6E76")
LOGO_W = 2.25 * 72

FLEX_FIELDS = ("Description", "Recommended Action", "Remarks", engine.LINK_FIELD)
WRAP_FIELDS = {"Responsible Party": 84.0, "Type": 60.0}   # short fields that may wrap, with a width cap
LINK_COL_W = 104.0                        # fits "https://www.planradar.com/dr/" per line
MIN_TEXT_W = 70.0                         # narrowest a text column gets (e.g. an empty Remarks)

_FONT_DIR = Path(__file__).resolve().parent / "assets" / "fonts"
_ARIAL = [("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
          ("/Library/Fonts/Arial.ttf", "/Library/Fonts/Arial Bold.ttf"),
          ("/usr/share/fonts/truetype/msttcorefonts/Arial.ttf",
           "/usr/share/fonts/truetype/msttcorefonts/Arial_Bold.ttf"),
          (_FONT_DIR / "Arial.ttf", _FONT_DIR / "Arial Bold.ttf")]
_FALLBACK = (_FONT_DIR / "LiberationSans-Regular.ttf", _FONT_DIR / "LiberationSans-Bold.ttf")


def _register_fonts() -> tuple[str, str]:
    if "LogPDF" in pdfmetrics.getRegisteredFontNames():
        return "LogPDF", "LogPDF-Bold"
    for regular, bold in [*_ARIAL, _FALLBACK]:
        if Path(regular).exists() and Path(bold).exists():
            pdfmetrics.registerFont(TTFont("LogPDF", str(regular)))
            pdfmetrics.registerFont(TTFont("LogPDF-Bold", str(bold)))
            pdfmetrics.registerFontFamily("LogPDF", normal="LogPDF", bold="LogPDF-Bold",
                                          italic="LogPDF", boldItalic="LogPDF-Bold")
            return "LogPDF", "LogPDF-Bold"
    return "Helvetica", "Helvetica-Bold"


def _rgb(fill) -> colors.Color | None:
    if fill is None or not getattr(fill, "fill_type", None):
        return None
    return colors.HexColor("#" + str(fill.fgColor.rgb)[-6:])


def _text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _as_date(value) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    try:
        return datetime.strptime(str(value).strip()[:10], "%Y-%m-%d").date()
    except (TypeError, ValueError):
        return None


def _log_sheet(wb, profile):
    title = engine.SHEET_TITLES.get(profile.key)
    if title in wb.sheetnames:
        return wb[title]
    return next(ws for ws in wb.worksheets if ws.title not in (engine.SUMMARY_SHEET, engine.AUDIT_SHEET))


def read_log(xlsx_bytes: bytes, profile) -> dict:
    """Pull the visible log out of the workbook: headers, rows, fills, alignment, title block."""
    wb = load_workbook(io.BytesIO(xlsx_bytes))
    ws = _log_sheet(wb, profile)
    n_cols = profile.last_col
    letters = [get_column_letter(c) for c in range(1, n_cols + 1)]
    headers = [_text(ws.cell(engine.HEADER_ROW, c).value) for c in range(1, n_cols + 1)]
    fields = {col: field for field, col in profile.column_map.items()}
    col_fields = [fields.get(letter, headers[i]) for i, letter in enumerate(letters)]
    widths = [ws.column_dimensions[l].width or 8.43 for l in letters]

    status_i = profile.status_col - 1
    ref = _as_date(ws["C5"].value) or date.today()
    obs_i = col_fields.index("Observation Date") if "Observation Date" in col_fields else None
    type_i = col_fields.index("Type") if "Type" in col_fields else None

    rows, row_fills, cell_fills = [], [], []
    last = ws.max_row
    for r in range(engine.DATA_START_ROW, last + 1):
        values = [ws.cell(r, c).value for c in range(1, n_cols + 1)]
        if all(v in (None, "") for v in values):
            continue
        if ws.row_dimensions[r].hidden:
            continue                                  # hidden by the open-items filter
        texts = [_text(v) for v in values]
        status = texts[status_i].lower()
        row_fill, status_fill = None, None
        if profile.age_highlight and obs_i is not None and status not in engine.CLOSED_WORDS:
            obs = _as_date(values[obs_i])
            if obs is not None:
                age = (ref - obs).days
                for lo, hi, _label, fill in engine.AGE_BANDS:
                    if age >= lo and (hi is None or age < hi):
                        row_fill = _rgb(fill)
        if profile.type_highlight and type_i is not None:
            for value, _name, fill in engine.TYPE_FILLS:
                if texts[type_i].upper() == value:
                    row_fill = _rgb(fill)
                    break
        if profile.status_colors:
            status_fill = {"open": _rgb(engine.OPEN_FILL), "in progress": _rgb(engine.INPROGRESS_FILL),
                           "closed": _rgb(engine.CLOSED_FILL)}.get(status)
        rows.append(texts)
        row_fills.append(row_fill)
        cell_fills.append(status_fill)

    legend_title, legend = None, []
    sw_col, lab_col = profile.legend_cols
    if profile.age_highlight or profile.type_highlight:
        legend_title = _text(ws[f"{sw_col}1"].value)
        for r in range(2, 7):
            label = _text(ws[f"{lab_col}{r}"].value)
            if label:
                legend.append((_rgb(ws[f"{sw_col}{r}"].fill), label))

    return {
        "sheet": ws.title, "project": _text(ws["A4"].value), "date_label": _text(ws["A5"].value),
        "date": _text(ws["C5"].value), "headers": headers, "fields": col_fields, "widths": widths,
        "rows": rows, "row_fills": row_fills, "cell_fills": cell_fills,
        "center": set(c - 1 for c in profile.center_columns), "status_i": status_i,
        "legend_title": legend_title, "legend": legend,
    }


def _column_widths(log: dict, font: str, bold: str) -> list[float]:
    """Short columns get exactly what their content needs. The link column gets a fixed
    width that fits the PlanRadar address on ~3 lines. The long-text columns share the
    rest by how much text each holds, so the busiest column gets the most room."""
    sw = pdfmetrics.stringWidth
    widths: list[float | None] = []
    for i, field in enumerate(log["fields"]):
        if field in FLEX_FIELDS and field != engine.LINK_FIELD:
            widths.append(None)
            continue
        words = log["headers"][i].split() or [""]
        head = max(sw(w, bold, BODY_SIZE) for w in words) + 2 * PAD_X + 1
        if field == engine.LINK_FIELD:
            widths.append(max(head, LINK_COL_W))
            continue
        need = head
        for row in log["rows"]:
            need = max(need, sw(row[i], font, BODY_SIZE) + 2 * PAD_X + 1)
        widths.append(min(need, max(head, WRAP_FIELDS.get(field, 1e9))))
    flex = [i for i, w in enumerate(widths) if w is None]
    left = CONTENT_W - sum(w for w in widths if w is not None)
    if not flex:
        scale = CONTENT_W / sum(widths)
        return [w * scale for w in widths]
    chars = {i: sum(len(row[i]) for row in log["rows"]) for i in flex}
    weights = {i: (chars[i] ** 0.5 if chars[i] else 0) for i in flex}
    floor = {i: max(MIN_TEXT_W, sw(log["headers"][i], bold, BODY_SIZE) + 2 * PAD_X + 1) for i in flex}
    spare = left - sum(floor.values())
    total = sum(weights.values())
    for i in flex:
        share = weights[i] / total if total else 1 / len(flex)
        widths[i] = floor[i] + spare * share
    return widths


def _styles(font: str, bold: str):
    base = dict(fontName=font, fontSize=BODY_SIZE, leading=LEADING, splitLongWords=1)
    return {
        "left": ParagraphStyle("left", alignment=TA_LEFT, **base),
        "center": ParagraphStyle("center", alignment=TA_CENTER, **base),
        "head": ParagraphStyle("head", alignment=TA_CENTER, **{**base, "fontName": bold}),
        "project": ParagraphStyle("project", fontName=bold, fontSize=11, leading=14),
        "meta": ParagraphStyle("meta", fontName=font, fontSize=9, leading=12),
        "legend_head": ParagraphStyle("lh", fontName=bold, fontSize=8, leading=10),
        "legend": ParagraphStyle("lg", fontName=font, fontSize=8, leading=10),
        "empty": ParagraphStyle("empty", alignment=TA_CENTER, textColor=MUTED, **base),
    }


def _para(text: str, style, link: bool = False) -> Paragraph:
    safe = escape(text).replace("\n", "<br/>")
    if link and text.lower().startswith(("http://", "https://")):
        safe = f'<a href="{escape(text, {chr(34): "&quot;"})}">{safe}</a>'
    return Paragraph(safe, style)


def _title_block(log: dict, st, open_count: int) -> Table:
    logo_path = engine.LOGO_PATH
    from PIL import Image as PILImage
    w_px, h_px = PILImage.open(logo_path).size
    logo = Image(str(logo_path), width=LOGO_W, height=LOGO_W * h_px / w_px, mask="auto")
    noun = "item" if open_count == 1 else "items"
    left = [logo, Spacer(1, 8),
            _para(log["project"], st["project"]),
            Paragraph(f'<font name="{st["legend_head"].fontName}">{escape(log["date_label"] or "Date Updated:")}</font> '
                      f'{escape(log["date"])}', st["meta"]),
            Paragraph(f'{escape(log["sheet"])} \u00b7 {open_count} open {noun} shown', st["meta"])]
    right = ""
    if log["legend"]:
        data = [[_para(log["legend_title"] or "Legend", st["legend_head"]), ""]]
        style = [("SPAN", (0, 0), (1, 0)), ("GRID", (0, 0), (-1, -1), THIN, colors.black),
                 ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                 ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                 ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]
        for r, (fill, label) in enumerate(log["legend"], start=1):
            data.append(["", _para(label, st["legend"])])
            if fill is not None:
                style.append(("BACKGROUND", (0, r), (0, r), fill))
        right = Table(data, colWidths=[22, 250], style=TableStyle(style), hAlign="RIGHT")
    block = Table([[left, right]], colWidths=[CONTENT_W - 290, 290],
                  style=TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                                    ("ALIGN", (1, 0), (1, 0), "RIGHT"),
                                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                    ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                                    ("TOPPADDING", (0, 0), (-1, -1), 0),
                                    ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
    return block


def _log_table(log: dict, st, widths: list[float]) -> Table:
    link_i = log["fields"].index(engine.LINK_FIELD) if engine.LINK_FIELD in log["fields"] else None
    data = [[_para(h, st["head"]) for h in log["headers"]]]
    for row in log["rows"]:
        data.append([_para(v, st["center"] if i in log["center"] else st["left"], link=i == link_i)
                     for i, v in enumerate(row)])
    style = [
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), PAD_X), ("RIGHTPADDING", (0, 0), (-1, -1), PAD_X),
        ("TOPPADDING", (0, 0), (-1, -1), PAD_Y), ("BOTTOMPADDING", (0, 0), (-1, -1), PAD_Y),
        ("GRID", (0, 1), (-1, -1), THIN, colors.black),
        ("GRID", (0, 0), (-1, 0), THICK, colors.black),
    ]
    if not log["rows"]:
        data.append([_para("No open items.", st["empty"])] + [""] * (len(widths) - 1))
        style += [("SPAN", (0, 1), (-1, 1))]
    for r, (row_fill, cell_fill) in enumerate(zip(log["row_fills"], log["cell_fills"]), start=1):
        if row_fill is not None:
            style.append(("BACKGROUND", (0, r), (-1, r), row_fill))
        if cell_fill is not None:
            style.append(("BACKGROUND", (log["status_i"], r), (log["status_i"], r), cell_fill))
    return Table(data, colWidths=widths, repeatRows=1, style=TableStyle(style), hAlign="LEFT")


def _numbered_canvas(footer_left: str, font: str):
    class NumberedCanvas(rl_canvas.Canvas):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self._pages = []

        def showPage(self):
            self._pages.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            total = len(self._pages)
            for state in self._pages:
                self.__dict__.update(state)
                self.setFont(font, 7.5)
                self.setFillColor(MUTED)
                base = MARGIN - pdfmetrics.getDescent(font, 7.5)    # text bottom sits on the margin
                self.drawString(MARGIN, base, footer_left)
                self.drawRightString(PAGE[0] - MARGIN, base, f"Page {self._pageNumber} of {total}")
                super().showPage()
            super().save()
    return NumberedCanvas


def build_log_pdf(xlsx_bytes: bytes, profile) -> bytes:
    """Landscape US Letter PDF of the log's visible (open) items."""
    font, bold = _register_fonts()
    log = read_log(xlsx_bytes, profile)
    st = _styles(font, bold)
    widths = _column_widths(log, font, bold)

    buf = io.BytesIO()
    doc = BaseDocTemplate(buf, pagesize=PAGE, leftMargin=MARGIN, rightMargin=MARGIN,
                          topMargin=MARGIN, bottomMargin=MARGIN + FOOTER_H,
                          title=f"{log['project']} \u2013 {log['sheet']}", author="Thornton Tomasetti")
    # No frame padding, so the table edges sit exactly on the 0.5 in margins.
    frame = Frame(MARGIN, MARGIN + FOOTER_H, CONTENT_W, PAGE[1] - 2 * MARGIN - FOOTER_H,
                  leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0, id="body")
    doc.addPageTemplates([PageTemplate(id="log", frames=[frame])])
    story = [_title_block(log, st, len(log["rows"])), Spacer(1, 12), _log_table(log, st, widths)]
    footer = f"{log['project']} \u00b7 {log['sheet']} \u00b7 {log['date_label'].rstrip(':') or 'Date Updated'} {log['date']}"
    doc.build(story, canvasmaker=_numbered_canvas(footer, font))
    return buf.getvalue()
