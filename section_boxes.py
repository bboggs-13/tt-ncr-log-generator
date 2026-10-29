"""Grey section boxes for TT site visit reports.

PlanRadar's raw export draws the Executive Summary, observation pages and appendix
as grey-bordered tables, but leaves the subsections of B. General Observations,
C. Follow-Up Items and D. Action Items as loose text. This module boxes each of those
subheadings (an 18 pt cell) and the content under it, in the same style as the rest of
the report: #DDDDDD borders, 0.7 pt wide, spanning 28.5-565.5 pt. Main headings
("B. General Observations") stay unboxed. Boxes close at a page break and reopen at
the top of the next page.

So the text clears the new borders the way it does in the finished TT template, text
inside a new box is mapped from 28.5-565.5 pt to 30-564 pt (a 1.5 pt inset with a
0.56% horizontal compression that isn't visible). This is done by editing the
page's own drawing commands, so text stays real, searchable text with no duplicates.
Subheadings that already sit in a grey box are left alone, so running this on a
finished report changes nothing.
"""
from __future__ import annotations

from dataclasses import dataclass, field
import re

import fitz

BORDER_GREY = (0xDD / 255,) * 3
LINE = 0.7                       # border width, pt
LEFT, RIGHT = 28.5, 565.5        # outer edges of the report's tables
TEXT_LEFT, TEXT_RIGHT = 30.0, 564.0
HEAD_ABOVE, HEAD_BELOW = 12.75, 5.25   # subheading cell, measured from the text baseline (18 pt cell)
CONTENT_PAD = 6.0                # space between the last line and the bottom border
CLEARANCE = 3.0                  # keep this far from whatever follows a box
ZERO_WIDTH = "\u200b\u200c\u200d\ufeff\u00a0"

_MAIN_HEADING = re.compile(r"^[A-Z]\.\s+\S")
_BOXED_SECTIONS = re.compile(r"^[BCD]\.\s", re.I)   # General Observations, Follow-Up Items, Action Items


@dataclass
class _Line:
    page: int
    rect: fitz.Rect
    text: str
    bold: bool
    size: float
    baseline: float

    @property
    def blank(self) -> bool:
        return not self.text


@dataclass
class BoxPlan:
    page_index: int
    rect: tuple[float, float, float, float]
    kind: str                      # "heading" or "content"
    title: str
    open_top: bool = False         # content continued from the previous page


@dataclass
class BoxResult:
    pdf: bytes
    boxes: list[BoxPlan] = field(default_factory=list)
    sections: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def _clean(text: str) -> str:
    return re.sub(f"[{ZERO_WIDTH}]", "", text).strip()


def _content_area(page: fitz.Page) -> tuple[float, float]:
    """Top and bottom of the body area (between the page header and the notice footer)."""
    top, bottom = 84.8, page.rect.height - 70.2
    for drawing in page.get_drawings():
        r = drawing["rect"]
        if drawing.get("fill") == (1.0, 1.0, 1.0) and r.x0 <= LEFT + 0.1 and r.width > 500 and r.height > 300:
            top, bottom = r.y0, r.y1
            break
    return top, bottom


def _grey_rules(page: fitz.Page) -> list[fitz.Rect]:
    rules = []
    for drawing in page.get_drawings():
        fill = drawing.get("fill")
        if fill and all(abs(c - 0.867) < 0.02 for c in fill[:3]):
            # A path can hold several edges (ours draw all four at once), so read each one.
            parts = [fitz.Rect(item[1]) for item in drawing["items"] if item[0] == "re"]
            rules.extend(parts or [fitz.Rect(drawing["rect"])])
    return rules


def _lines(page: fitz.Page, index: int, top: float, bottom: float) -> list[_Line]:
    out = []
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            spans = line["spans"]
            rect = fitz.Rect(line["bbox"])
            if not spans or rect.y0 < top - 1 or rect.y1 > bottom + 1:
                continue
            text = _clean("".join(s["text"] for s in spans))
            first = next((s for s in spans if _clean(s["text"])), spans[0])
            out.append(_Line(index, rect, text, "Bold" in first["font"] or bool(first["flags"] & 16),
                             round(first["size"], 2), first["origin"][1]))
    out.sort(key=lambda l: (l.rect.y0, l.rect.x0))
    return out


def _inside_existing_box(line: _Line, rules: list[fitz.Rect]) -> bool:
    """True when grey rules sit just above and below the line (it's already in a table)."""
    above = any(r.y1 <= line.rect.y0 + 1 and line.rect.y0 - r.y1 < 25 and r.width > 100
                and r.x0 <= line.rect.x0 + 2 for r in rules)
    below = any(r.y0 >= line.rect.y1 - 1 and r.y0 - line.rect.y1 < 40 and r.width > 100
                and r.x0 <= line.rect.x0 + 2 for r in rules)
    sides = any(r.width < 2 and r.x0 <= line.rect.x0 and r.y0 <= line.rect.y0 and r.y1 >= line.rect.y1
                for r in rules)
    return (above and below) or sides


def _is_subheading(line: _Line) -> bool:
    return (line.bold and line.rect.x0 <= LEFT + 3.5 and line.text.endswith(":")
            and not _MAIN_HEADING.match(line.text) and 8 <= line.size <= 10)


def plan_boxes(doc: fitz.Document) -> tuple[list[BoxPlan], list[str]]:
    """Find the unboxed subsections of B/C/D and lay out their boxes page by page."""
    flow: list[_Line] = []
    areas, rules = {}, {}
    for index, page in enumerate(doc):
        areas[index] = _content_area(page)
        rules[index] = _grey_rules(page)
        flow.extend(_lines(page, index, *areas[index]))

    plans: list[BoxPlan] = []
    titles: list[str] = []
    in_boxed_section = False
    i = 0
    while i < len(flow):
        line = flow[i]
        if line.bold and _MAIN_HEADING.match(line.text) and line.size >= 10:
            in_boxed_section = bool(_BOXED_SECTIONS.match(line.text))
            i += 1
            continue
        if not (in_boxed_section and _is_subheading(line)) or _inside_existing_box(line, rules[line.page]):
            i += 1
            continue

        # The subheading cell.
        head = fitz.Rect(LEFT, line.baseline - HEAD_ABOVE, RIGHT, line.baseline + HEAD_BELOW)
        plans.append(BoxPlan(line.page, tuple(head), "heading", line.text))
        titles.append(line.text)

        # Content runs until the next subheading, main heading, existing table or page end.
        j = i + 1
        body: list[_Line] = []
        stop: _Line | None = None
        while j < len(flow):
            nxt = flow[j]
            if (nxt.bold and _MAIN_HEADING.match(nxt.text) and nxt.size >= 10) or _is_subheading(nxt) \
                    or (not nxt.blank and _inside_existing_box(nxt, rules[nxt.page])):
                stop = nxt
                break
            body.append(nxt)
            j += 1

        next_is_sibling = stop is not None and _is_subheading(stop) and not _inside_existing_box(stop, rules[stop.page])
        pages = sorted({l.page for l in body if not l.blank} | {line.page})
        for page_index in pages:
            top_area, bottom_area = areas[page_index]
            on_page = [l for l in body if l.page == page_index and not l.blank]
            continued_from_before = page_index != line.page
            top = top_area if continued_from_before else head.y1
            last_page = page_index == pages[-1]
            if on_page:
                bottom = max(l.rect.y1 for l in on_page) + CONTENT_PAD
            else:
                bottom = top + CONTENT_PAD
            if last_page and stop is not None and stop.page == page_index:
                if next_is_sibling:
                    bottom = stop.baseline - HEAD_ABOVE          # meet the next subheading cell
                else:
                    bottom = min(bottom, stop.rect.y0 - CLEARANCE)
            bottom = min(bottom, bottom_area - 1.0)
            if bottom - top > 2:
                plans.append(BoxPlan(page_index, (LEFT, top, RIGHT, bottom), "content", line.text,
                                     open_top=continued_from_before))
        i = j
    return plans, titles


# --- Content stream editing ------------------------------------------------------
_DELIM = b"()<>[]{}/%"
_WS = b" \t\r\n\f\x00"


def _tokens(data: bytes):
    """Yield (kind, value, start, end) for a PDF content stream. Kinds: op, num, other."""
    i, n = 0, len(data)
    while i < n:
        c = data[i]
        if c in _WS:
            i += 1
        elif c == 0x25:                                        # % comment
            while i < n and data[i] not in b"\r\n":
                i += 1
        elif c == 0x28:                                        # (string)
            start, depth = i, 0
            while i < n:
                ch = data[i]
                if ch == 0x5C:
                    i += 2
                    continue
                if ch == 0x28:
                    depth += 1
                elif ch == 0x29:
                    depth -= 1
                    if depth == 0:
                        i += 1
                        break
                i += 1
            yield "other", None, start, i
        elif c == 0x3C and data[i:i + 2] != b"<<":             # <hex>
            start = i
            i = data.index(b">", i) + 1
            yield "other", None, start, i
        elif data[i:i + 2] in (b"<<", b">>"):
            yield "other", None, i, i + 2
            i += 2
        elif c in b"[]{}":
            yield "other", None, i, i + 1
            i += 1
        elif c == 0x2F:                                        # /Name
            start = i
            i += 1
            while i < n and data[i] not in _WS and data[i] not in _DELIM:
                i += 1
            yield "other", None, start, i
        else:
            start = i
            while i < n and data[i] not in _WS and data[i] not in _DELIM:
                i += 1
            word = data[start:i]
            try:
                yield "num", float(word), start, i
            except ValueError:
                yield "op", word, start, i
                if word == b"ID":                              # inline image data
                    m = re.compile(rb"[\s]EI(?=[\s]|$)").search(data, i + 1)
                    i = m.end() if m else n


def _fmt(m: fitz.Matrix) -> bytes:
    return (" ".join(f"{v:.6f}".rstrip("0").rstrip(".") or "0" for v in (m.a, m.b, m.c, m.d, m.e, m.f))
            + " cm").encode()


def _shift_text(page: fitz.Page, bands: list[tuple[float, float]]) -> int:
    """Map text whose baseline falls in a band from 28.5-565.5 to 30-564 pt. Returns runs moved."""
    doc = page.parent
    data = b"\n".join(doc.xref_stream(x) or b"" for x in page.get_contents())
    to_page = page.transformation_matrix                    # PDF space -> PyMuPDF page space
    from_page = ~to_page
    scale = (TEXT_RIGHT - TEXT_LEFT) / (RIGHT - LEFT)
    squeeze = fitz.Matrix(scale, 0, 0, 1, TEXT_LEFT - LEFT * scale, 0)   # in page space
    squeeze_pdf = to_page * squeeze * from_page

    ctm, stack = fitz.Matrix(1, 0, 0, 1, 0, 0), []
    operands: list[float] = []
    edits: list[tuple[int, bytes]] = []
    bt_start, bt_ctm, tlm, leading, hit = None, None, fitz.Matrix(1, 0, 0, 1, 0, 0), 0.0, False

    def baseline_y(matrix: fitz.Matrix) -> float:
        return fitz.Point(0, 0).transform(matrix * bt_ctm * to_page).y

    for kind, value, start, end in _tokens(data):
        if kind == "num":
            operands.append(value)
            continue
        if kind == "other":
            continue
        op = value
        if op == b"q":
            stack.append(ctm)
        elif op == b"Q":
            ctm = stack.pop() if stack else ctm
        elif op == b"cm" and len(operands) >= 6:
            ctm = fitz.Matrix(*operands[-6:]) * ctm
        elif op == b"BT":
            bt_start, bt_ctm, tlm, hit = start, ctm, fitz.Matrix(1, 0, 0, 1, 0, 0), False
        elif bt_start is not None and op == b"Tm" and len(operands) >= 6:
            tlm = fitz.Matrix(*operands[-6:])
        elif bt_start is not None and op in (b"Td", b"TD") and len(operands) >= 2:
            tlm = fitz.Matrix(1, 0, 0, 1, operands[-2], operands[-1]) * tlm
            if op == b"TD":
                leading = -operands[-1]
        elif op == b"TL" and operands:
            leading = operands[-1]
        elif bt_start is not None and op in (b"T*", b"'", b'"'):
            tlm = fitz.Matrix(1, 0, 0, 1, 0, -leading) * tlm
        if bt_start is not None and op in (b"Tj", b"TJ", b"'", b'"'):
            y = baseline_y(tlm)
            if any(top <= y <= bottom for top, bottom in bands):
                hit = True
        if op == b"ET" and bt_start is not None:
            if hit:
                local = bt_ctm * squeeze_pdf * ~bt_ctm
                edits.append((bt_start, b"q " + _fmt(local) + b"\n"))
                edits.append((end, b"\nQ"))
            bt_start = None
        operands.clear()

    if not edits:
        return 0
    pieces, last = [], 0
    for pos, insert in sorted(edits, key=lambda e: e[0]):
        pieces.append(data[last:pos])
        pieces.append(insert)
        last = pos
    pieces.append(data[last:])
    xref = doc.get_new_xref()
    doc.update_object(xref, "<<>>")
    doc.update_stream(xref, b"".join(pieces))
    page.set_contents(xref)
    return len(edits) // 2


def _draw_box(page: fitz.Page, rect: fitz.Rect, top_edge: bool = True) -> None:
    edges = [
        fitz.Rect(rect.x0, rect.y1, rect.x1, rect.y1 + LINE),                 # bottom
        fitz.Rect(rect.x0, rect.y0, rect.x0 + LINE, rect.y1 + LINE),          # left
        fitz.Rect(rect.x1 - LINE, rect.y0, rect.x1, rect.y1 + LINE),          # right
    ]
    if top_edge:
        edges.insert(0, fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y0 + LINE))
    shape = page.new_shape()
    for edge in edges:
        shape.draw_rect(edge)
    shape.finish(color=None, fill=BORDER_GREY, width=0)
    shape.commit(overlay=True)


def apply_section_boxes(pdf_bytes: bytes) -> BoxResult:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        plans, titles = plan_boxes(doc)
        if not plans:
            return BoxResult(pdf_bytes, [], [], [])
        warnings = []
        for page_index in sorted({p.page_index for p in plans}):
            page = doc[page_index]
            bands = [(p.rect[1] - 0.5, p.rect[3] + 0.5) for p in plans if p.page_index == page_index]
            _shift_text(page, bands)
            for plan in (p for p in plans if p.page_index == page_index):
                _draw_box(page, fitz.Rect(plan.rect))
        output = doc.tobytes(garbage=3, deflate=True)
    finally:
        doc.close()
    return BoxResult(output, plans, titles, warnings)
