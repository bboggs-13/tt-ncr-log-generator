"""Grey section boxes for TT site visit reports.

PlanRadar's raw export draws the Executive Summary, observation pages and appendix
as grey-bordered tables, but leaves the subsections of General Observations,
Follow-Up Items, Action Items and Summary of Observations as loose text (whatever
letter each section has in the report format). This module boxes each of those
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

# Spacing clean-up ("tighten spacing"). The standard gap is measured from the report itself
# (median gap between list items in the summary sections, 22.2 pt in PlanRadar's export).
DEFAULT_ITEM_GAP = 22.2          # used when the report has too few items to measure
DEFAULT_HEAD_GAP = 6.0           # subheading -> first line of its content
TIGHTEN_OVER = 6.0               # gaps more than this over the standard are tightened
EVEN_UNDER = 6.0                 # list items closer than standard minus this are spaced out (if room)
ROOM_MARGIN = 2.0                # moved content must stay this far above the page's body area
_ITEM_START = re.compile(r"^(\d+\.\s|\d+\.[A-Z]{2,4}\.\d+|Note:)")

# Table cells (Appendix, Executive Summary, observation pages) with empty space at the bottom.
DEFAULT_CELL_PAD = 3.2           # last text line -> bottom border, as in the report's own cells
DEFAULT_CELL_TOP = 4.4           # top border -> first text line
TRIM_OVER = 8.0                  # trim a cell when it has more than this much extra space
GRID = 0.5                       # moves are whole multiples of this (keeps the render check exact)
INK = 235                        # grey level (0-255) darker than this counts as content
_NO_TRIM = re.compile(r"\b(Signed|Signature)\b", re.I)   # signature cells are meant to be empty

_MAIN_HEADING = re.compile(r"^[A-Z]\.\s+\S")
# Sections are matched by name, not letter, because the letters move between report
# formats (the newer format puts B. Follow-Up Items first and the observation summary
# in E. Summary of Observations, after the ticket pages). Ticket pages such as
# "D. General Observation #1" (singular) are not matched.
_BOXED_SECTIONS = re.compile(
    r"^[A-Z]\.\s+(General Observations|Follow[-\s]?Up Items|Action Items|Summary of Observations)\b", re.I)


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
class Move:
    """Close up (dy < 0) or open (dy > 0) space on one page.

    dy < 0: the band y0-y1 is removed; content below it moves up by -dy and anything
    crossing the band (cell borders, fills, clipping) gets shorter.
    dy > 0: content from y0 down moves down by dy."""
    page_index: int
    y0: float
    y1: float
    dy: float
    old_gap: float
    new_gap: float
    label: str
    kind: str = "gap"              # "gap" (summary sections) or "cell" (existing table cells)

    def note(self) -> str:
        page = f"Page {self.page_index + 1}"
        if self.kind == "cell":
            return f"{page}: trimmed {abs(self.dy):.1f} pt of empty space at the bottom of {self.label}."
        if self.kind == "cell_top":
            return f"{page}: trimmed {abs(self.dy):.1f} pt of empty space at the top of {self.label}."
        verb = "tightened" if self.dy < 0 else "evened out"
        return f"{page}: {verb} the gap above \u201c{self.label}\u201d from {self.old_gap:.1f} pt to {self.new_gap:.1f} pt."


@dataclass
class _Section:
    head: _Line
    body: list[_Line]
    stop: "_Line | None"
    next_is_sibling: bool


@dataclass
class LayoutPlan:
    boxes: list[BoxPlan]
    titles: list[str]
    moves: list[Move] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    skipped: frozenset = frozenset()


@dataclass
class BoxResult:
    pdf: bytes
    boxes: list[BoxPlan] = field(default_factory=list)
    sections: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    moves: list[Move] = field(default_factory=list)


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


def _find_sections(flow: list[_Line], rules: dict) -> list[_Section]:
    """Unboxed subsections of the summary sections, each with its content and what stops it."""
    sections: list[_Section] = []
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
        sibling = stop is not None and _is_subheading(stop) and not _inside_existing_box(stop, rules[stop.page])
        sections.append(_Section(line, body, stop, sibling))
        i = j
    return sections


def _median(values: list[float], default: float) -> float:
    values = sorted(values)
    if not values:
        return default
    mid = len(values) // 2
    return round(values[mid] if len(values) % 2 else (values[mid - 1] + values[mid]) / 2, 1)


def _label(text: str) -> str:
    words = text.split()
    return " ".join(words[:6]) + ("\u2026" if len(words) > 6 else "")


def _moves_for(sections: list[_Section], boxed: bool = True) -> tuple[list[Move], float]:
    """Gaps to tighten (too big) or even out (list items with no gap), per page."""
    item_gaps, head_gaps = [], []
    for sec in sections:
        seq = [l for l in sec.body if not l.blank]
        if seq and seq[0].page == sec.head.page:
            head_gaps.append(seq[0].rect.y0 - sec.head.rect.y1)
        for a, b in zip(seq, seq[1:]):
            gap = b.rect.y0 - a.rect.y1
            if a.page == b.page and 15 <= gap <= 30:
                item_gaps.append(gap)
    item_gap = _median(item_gaps, DEFAULT_ITEM_GAP)
    head_gap = _median([g for g in head_gaps if 0 <= g <= 15], DEFAULT_HEAD_GAP)

    moves: dict[tuple[int, float], Move] = {}
    for sec in sections:
        seq = [sec.head] + [l for l in sec.body if not l.blank]
        ends_at_heading = sec.stop is not None and (
            sec.next_is_sibling or (sec.stop.bold and _MAIN_HEADING.match(sec.stop.text) and sec.stop.size >= 10))
        if ends_at_heading:
            seq.append(sec.stop)
        for a, b in zip(seq, seq[1:]):
            if a.page != b.page:
                continue
            gap = b.rect.y0 - a.rect.y1
            target = head_gap if a is sec.head else item_gap
            if boxed and b is sec.stop and sec.next_is_sibling:
                # Close the box CONTENT_PAD below the last line, like every other box,
                # and start the next subheading cell right there.
                target = CONTENT_PAD + HEAD_ABOVE - (b.baseline - b.rect.y0)
            limit = head_gap + 10 if a is sec.head else max(target, item_gap) + TIGHTEN_OVER
            if boxed and b is sec.stop and sec.next_is_sibling:
                limit = target + 2.0                            # every box closes the same way
            if gap > limit:
                pass                                            # too big: tighten
            elif (a is not sec.head and b is not sec.stop and _ITEM_START.match(b.text)
                  and -1 < gap < item_gap - EVEN_UNDER):
                pass                                            # a list item with no gap: even out
            else:
                continue
            dy = _q(target - gap)
            if not dy:
                continue
            if dy < 0:                                          # remove a band just above line b
                y1 = _q(b.rect.y0, "down")
                y0 = y1 + dy
                if y0 < a.rect.y1 - 0.01:
                    continue
            else:                                               # open space between a and b
                y0 = y1 = _q((a.rect.y1 + b.rect.y0) / 2)
            key = (b.page, round(b.baseline, 1))
            moves[key] = Move(b.page, y0, y1, dy, round(gap, 1), round(gap + dy, 1), _label(b.text))
    return sorted(moves.values(), key=lambda m: (m.page_index, m.y0)), item_gap


def _q(v: float, mode: str = "near") -> float:
    """Snap to the GRID (moves by whole render pixels, so the render check can be exact)."""
    import math
    f = {"near": round, "up": math.ceil, "down": math.floor}[mode]
    return f(v / GRID) * GRID


def _shift(moves: list[Move], y: float) -> float:
    """How far a point at height y moves."""
    total = 0.0
    for m in moves:
        if m.dy < 0:
            if y >= m.y1:
                total += m.dy
            elif y > m.y0:
                total -= y - m.y0
        elif y >= m.y0:
            total += m.dy
    return total


def _render_gray(page: fitz.Page, zoom: float = 1 / GRID):
    import numpy as np
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), colorspace=fitz.csGRAY, alpha=False)
    return np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.stride)[:, :pix.width]


def _is_background(rect: fitz.Rect, area: tuple[float, float]) -> bool:
    return abs(rect.y0 - area[0]) < 0.6 and abs(rect.y1 - area[1]) < 0.6 and rect.width > 500


@dataclass
class _Row:
    page: int
    # (area_top is filled in by _table_rows; used to skip rows continued from the page before)
    top: float                     # top border y
    bottom: float                  # bottom border y
    text_bottom: float | None
    ink_bottom: float | None
    label: str
    text_top: float | None = None
    ink_top: float | None = None
    area_top: float = 0.0


def _table_rows(page: fitz.Page, index: int, area: tuple[float, float], lines: list[_Line]) -> list[_Row]:
    """Full-width table rows (grey border above and below, left border along the side)."""
    import numpy as np
    top_area, bottom_area = area
    rules = _grey_rules(page)
    horiz = [r for r in rules if r.height < 2 and r.width > 3 * r.height]
    vert = [r for r in rules if r.width < 2 and r.height > 3 * r.width]
    ys = sorted({round(r.y0 * 4) / 4 for r in horiz if top_area - 0.5 <= r.y0 <= bottom_area + 0.5})

    def covers(y: float) -> bool:
        spans = sorted((r.x0, r.x1) for r in horiz if abs(r.y0 - y) < 0.3)
        reach = LEFT + 1
        for x0, x1 in spans:
            if x0 <= reach + 1:
                reach = max(reach, x1)
        return reach >= RIGHT - 1

    rows: list[_Row] = []
    gray = None
    prev_label = ""
    for y_top, y_bot in zip(ys, ys[1:]):
        if not (covers(y_top) and covers(y_bot)):
            continue
        if not any(v.x0 < LEFT + 1 and v.y0 <= y_top + 1 and v.y1 >= y_bot - 1 for v in vert):
            continue
        in_row = [l for l in lines if not l.blank and l.rect.y0 >= y_top - 1 and l.rect.y1 <= y_bot + 1]
        label = in_row[0].text if in_row else ""
        if gray is None:
            gray = _render_gray(page)
        z = 1 / GRID
        r0, r1 = int((y_top + LINE + 1.0) * z), int(y_bot * z)
        c0, c1 = int((LEFT + LINE + 0.5) * z), int((RIGHT - LINE - 0.5) * z)
        block = gray[r0:r1, c0:c1].copy()
        for v in vert:                                         # column dividers aren't content
            if v.y0 <= y_top + 1 and v.y1 >= y_bot - 1:
                a, b = int((v.x0 - 1) * z) - c0, int((v.x1 + 1) * z) - c0
                block[:, max(0, a):max(0, b)] = 255
        ink_rows = np.where((block < INK).any(axis=1))[0]
        ink = (r0 + ink_rows.max() + 1) / z if len(ink_rows) else None
        ink_top = (r0 + ink_rows.min()) / z if len(ink_rows) else None
        rows.append(_Row(index, y_top, y_bot, max((l.rect.y1 for l in in_row), default=None), ink,
                         label or prev_label, min((l.rect.y0 for l in in_row), default=None), ink_top, top_area))
        prev_label = label or prev_label
    return rows


def _content_top_of(row: "_Row") -> float:
    return row.area_top


def _cell_trims(rows: list[_Row], pad: float, lines_by_page: dict, top_pad: float = DEFAULT_CELL_TOP) -> list[Move]:
    moves = []
    for row in rows:
        if any(_NO_TRIM.search(l.text) for l in lines_by_page.get(row.page, [])
               if not l.blank and l.rect.y0 >= row.top - 1 and l.rect.y1 <= row.bottom + 1):
            continue
        # Empty space at the top of the cell (text only: images sit right on the border).
        at_page_top = row.top <= _content_top_of(row) + 1.0      # continues from the page before
        if not at_page_top and row.text_top is not None and (row.ink_top is None or row.ink_top >= row.text_top - 1):
            y0 = _q(row.top + LINE, "up")
            y1 = _q(row.text_top - top_pad + (y0 - row.top), "down")
            if y1 - y0 >= TRIM_OVER:
                moves.append(Move(row.page, y0, y1, -(y1 - y0), round(row.text_top - row.top, 1),
                                  round(row.text_top - row.top - (y1 - y0), 1),
                                  f"the \u201c{_label(row.label.rstrip(':'))}\u201d cell", kind="cell_top"))
        ends = []
        if row.text_bottom is not None:
            ends.append(row.text_bottom + pad)
        image_row = row.ink_bottom is not None and (row.text_bottom is None or row.ink_bottom > row.text_bottom + 1)
        if image_row:
            ends.append(row.ink_bottom)                        # plans sit right on the border
        if not ends:
            continue                                           # an empty cell is left alone
        y0, y1 = _q(max(ends), "up"), _q(row.bottom, "down")
        if y1 - y0 < TRIM_OVER:
            continue
        if image_row and row.text_bottom is None:
            label = "the location plan cell" if re.search(r"Floor|Location|Elevation|Grid", row.label) else "an image cell"
        else:
            label = f"the \u201c{_label(row.label.rstrip(':'))}\u201d cell"
        moves.append(Move(row.page, y0, y1, -(y1 - y0), round(row.bottom - (max(ends) - (0 if image_row else pad)), 1),
                          round(y0 - (max(ends) - (0 if image_row else pad)), 1), label, kind="cell"))
    return moves


_UNMOVABLE = re.compile(rb"(?<![A-Za-z*'\"])(BI|sh)(?![A-Za-z*'\"])")


def _check_page_moves(page: fitz.Page, moves: list[Move], area: tuple[float, float],
                      lines: list[_Line]) -> tuple[list[Move], str | None]:
    """Drop moves that aren't safe, or that would push content past the page."""
    if not moves:
        return moves, None
    top_area, bottom_area = area
    stream = b"\n".join(page.parent.xref_stream(x) or b"" for x in page.get_contents())
    if _UNMOVABLE.search(stream):
        return [], f"Page {page.number + 1}: spacing left as-is (the page has content that can\u2019t be moved safely)."

    def overflows(candidate: list[Move]) -> bool:
        for line in lines:
            if not line.blank and line.rect.y1 + _shift(candidate, line.baseline) > bottom_area - ROOM_MARGIN:
                return True
        for drawing in page.get_drawings():
            r = drawing["rect"]
            if r.y0 >= top_area - 0.5 and r.y1 <= bottom_area + 0.5 and not _is_background(r, area) and \
                    r.y1 + _shift(candidate, r.y1) > bottom_area + 0.5:
                return True
        return False

    if overflows(moves):
        kept = [m for m in moves if m.dy < 0]
        return kept, f"Page {page.number + 1}: not enough room to even out list spacing, so only large gaps were tightened."
    return moves, None


def plan_layout(doc: fitz.Document, tighten: bool = False, skip_pages: frozenset = frozenset(),
                boxed: bool = True, protect_pages: frozenset = frozenset()) -> LayoutPlan:
    """Find the unboxed subsections of the summary sections, optionally even out their
    spacing, and lay out their boxes page by page (positions are after any moves)."""
    flow: list[_Line] = []
    areas, rules = {}, {}
    for index, page in enumerate(doc):
        areas[index] = _content_area(page)
        rules[index] = _grey_rules(page)
        flow.extend(_lines(page, index, *areas[index]))

    sections = _find_sections(flow, rules)
    moves: list[Move] = []
    notes: list[str] = []
    if tighten:
        planned = _moves_for(sections, boxed)[0] if sections else []
        lines_by_page: dict[int, list[_Line]] = {}
        for line in flow:
            lines_by_page.setdefault(line.page, []).append(line)
        rows = []
        for index, page in enumerate(doc):
            if index not in protect_pages and index not in skip_pages:
                rows.extend(_table_rows(page, index, areas[index], lines_by_page.get(index, [])))
        pads = [r.bottom - r.text_bottom for r in rows if r.text_bottom is not None
                and (r.ink_bottom is None or r.ink_bottom <= r.text_bottom + 1) and 1 <= r.bottom - r.text_bottom <= 8]
        tops = [r.text_top - r.top for r in rows if r.text_top is not None and 1 <= r.text_top - r.top <= 8]
        planned += _cell_trims(rows, _median(pads, DEFAULT_CELL_PAD), lines_by_page, _median(tops, DEFAULT_CELL_TOP))
        for page_index in sorted({m.page_index for m in planned}):
            page_moves = sorted((m for m in planned if m.page_index == page_index), key=lambda m: m.y0)
            if page_index in protect_pages:
                continue                                       # photo pages: captions use their positions
            if page_index in skip_pages:
                notes.append(f"Page {page_index + 1}: spacing left as-is (it couldn\u2019t be moved without "
                             f"changing how the page looks).")
                continue
            kept, note = _check_page_moves(doc[page_index], page_moves, areas[page_index],
                                           lines_by_page.get(page_index, []))
            moves.extend(kept)
            if note:
                notes.append(note)
        # Shift the lines so the boxes are laid out where the text ends up.
        for line in flow:
            dy = _shift([m for m in moves if m.page_index == line.page], line.baseline)
            if dy:
                line.rect = fitz.Rect(line.rect.x0, line.rect.y0 + dy, line.rect.x1, line.rect.y1 + dy)
                line.baseline += dy
        notes = [m.note() for m in moves] + notes

    plans: list[BoxPlan] = []
    titles: list[str] = []
    for sec in sections:
        line, body, stop = sec.head, sec.body, sec.stop
        head = fitz.Rect(LEFT, line.baseline - HEAD_ABOVE, RIGHT, line.baseline + HEAD_BELOW)
        plans.append(BoxPlan(line.page, tuple(head), "heading", line.text))
        titles.append(line.text)
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
                if sec.next_is_sibling:
                    bottom = stop.baseline - HEAD_ABOVE          # meet the next subheading cell
                else:
                    bottom = min(bottom, stop.rect.y0 - CLEARANCE)
            bottom = min(bottom, bottom_area - 1.0)
            if bottom - top > 2:
                plans.append(BoxPlan(page_index, (LEFT, top, RIGHT, bottom), "content", line.text,
                                     open_top=continued_from_before))
    return LayoutPlan(plans, titles, moves, notes, skip_pages)


def plan_boxes(doc: fitz.Document, tighten: bool = False) -> tuple[list[BoxPlan], list[str]]:
    """Boxes and subsection titles (kept for callers that only need the preview)."""
    plan = plan_layout(doc, tighten=tighten)
    return plan.boxes, plan.titles


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


class _Conflict(Exception):
    """A drawing unit can't be moved as one piece (e.g. one text block spans a gap)."""


_PATH_BUILD = {b"m", b"l", b"c", b"v", b"y", b"re", b"h"}
_PATH_PAINT = {b"S", b"s", b"f", b"F", b"f*", b"B", b"B*", b"b", b"b*", b"n"}


def _num(v: float) -> bytes:
    return (f"{v:.4f}".rstrip("0").rstrip(".") or "0").encode()


def _rewrite_page(page: fitz.Page, bands: list[tuple[float, float]], moves: list[Move],
                  area: tuple[float, float]) -> int:
    """Inset text in the boxed bands (28.5-565.5 -> 30-564 pt) and apply the moves.

    Text blocks and images are wrapped in their own transform. Path coordinates are
    rewritten point by point, so a border or clip that crosses a removed band gets
    shorter while everything below it moves up. Raises _Conflict if something can't be
    moved whole. Returns the number of changed drawing units."""
    doc = page.parent
    data = b"\n".join(doc.xref_stream(x) or b"" for x in page.get_contents())
    to_page = page.transformation_matrix                    # PDF space -> PyMuPDF page space
    from_page = ~to_page
    scale = (TEXT_RIGHT - TEXT_LEFT) / (RIGHT - LEFT)
    squeeze = fitz.Matrix(scale, 0, 0, 1, TEXT_LEFT - LEFT * scale, 0)   # in page space
    top_area, bottom_area = area

    ctm, stack = fitz.Matrix(1, 0, 0, 1, 0, 0), []
    operands: list[tuple[float, int, int]] = []
    edits: list[tuple[int, int, bytes]] = []                 # (start, end, replacement)
    bt_start, bt_ctm, tlm, leading = None, None, fitz.Matrix(1, 0, 0, 1, 0, 0), 0.0
    bt_hit, bt_dys = False, set()
    path: list[tuple[bytes, list[tuple[float, int, int]]]] = []
    path_ctm = None
    last_name: tuple[int, int] | None = None
    changed = 0

    def baseline_y(matrix: fitz.Matrix) -> float:
        return fitz.Point(0, 0).transform(matrix * bt_ctm * to_page).y

    def in_body(rect: fitz.Rect) -> bool:
        return rect.y0 >= top_area - 0.6 and rect.y1 <= bottom_area + 0.6

    def wrap(matrix_ctm: fitz.Matrix, page_matrix: fitz.Matrix) -> bytes:
        local = matrix_ctm * (to_page * page_matrix * from_page) * ~matrix_ctm
        return b"q " + _fmt(local) + b"\n"

    for kind, value, start, end in _tokens(data):
        if kind == "num":
            operands.append((value, start, end))
            continue
        if kind == "other":
            if data[start:start + 1] == b"/":
                last_name = (start, end)
            continue
        op = value
        vals = [o[0] for o in operands]
        if op == b"q":
            stack.append(ctm)
        elif op == b"Q":
            ctm = stack.pop() if stack else ctm
        elif op == b"cm" and len(vals) >= 6:
            ctm = fitz.Matrix(*vals[-6:]) * ctm
        elif op == b"BT":
            bt_start, bt_ctm, tlm, bt_hit, bt_dys = start, ctm, fitz.Matrix(1, 0, 0, 1, 0, 0), False, set()
        elif bt_start is not None and op == b"Tm" and len(vals) >= 6:
            tlm = fitz.Matrix(*vals[-6:])
        elif bt_start is not None and op in (b"Td", b"TD") and len(vals) >= 2:
            tlm = fitz.Matrix(1, 0, 0, 1, vals[-2], vals[-1]) * tlm
            if op == b"TD":
                leading = -vals[-1]
        elif op == b"TL" and vals:
            leading = vals[-1]
        elif bt_start is not None and op in (b"T*", b"'", b'"'):
            tlm = fitz.Matrix(1, 0, 0, 1, 0, -leading) * tlm
        elif bt_start is None and op in _PATH_BUILD:
            if not path:
                path_ctm = ctm
            path.append((op, list(operands)))
        elif bt_start is None and op in _PATH_PAINT and path:
            if moves:
                m = path_ctm * to_page
                inv = ~m
                pts = []
                for pop, ops in path:
                    nums = [o[0] for o in ops]
                    if pop == b"re" and len(nums) >= 4:
                        x, y, w, h = nums[-4:]
                        pts += [fitz.Point(x, y) * m, fitz.Point(x + w, y + h) * m]
                    elif pop != b"h":
                        pts += [fitz.Point(nums[k], nums[k + 1]) * m for k in range(0, len(nums) - 1, 2)]
                rect = fitz.Rect(min(p.x for p in pts), min(p.y for p in pts),
                                 max(p.x for p in pts), max(p.y for p in pts)) if pts else None
                if rect is not None and in_body(rect) and not _is_background(rect, area) \
                        and any(_shift(moves, p.y) for p in pts):
                    for pop, ops in path:
                        nums = [o[0] for o in ops]
                        if pop == b"re" and len(ops) >= 4:
                            x, y, w, h = nums[-4:]
                            a = fitz.Point(x, y) * m
                            b = fitz.Point(x + w, y + h) * m
                            a2 = fitz.Point(a.x, a.y + _shift(moves, a.y)) * inv
                            b2 = fitz.Point(b.x, b.y + _shift(moves, b.y)) * inv
                            new = (a2.x, a2.y, b2.x - a2.x, b2.y - a2.y)
                            for (v, s0, e0), nv in zip(ops[-4:], new):
                                if abs(nv - v) > 1e-6:
                                    edits.append((s0, e0, _num(nv)))
                        elif pop != b"h":
                            for k in range(0, len(ops) - 1, 2):
                                pt = fitz.Point(nums[k], nums[k + 1]) * m
                                new = fitz.Point(pt.x, pt.y + _shift(moves, pt.y)) * inv
                                for (v, s0, e0), nv in ((ops[k], new.x), (ops[k + 1], new.y)):
                                    if abs(nv - v) > 1e-6:
                                        edits.append((s0, e0, _num(nv)))
                    changed += 1
            path = []
        elif op == b"Do" and moves and bt_start is None:
            corners = [fitz.Point(x, y) * (ctm * to_page) for x in (0, 1) for y in (0, 1)]
            rect = fitz.Rect(min(p.x for p in corners), min(p.y for p in corners),
                             max(p.x for p in corners), max(p.y for p in corners))
            dy = _shift(moves, rect.y0) if in_body(fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y0)) else 0.0
            if dy:
                if abs(_shift(moves, rect.y1) - dy) > 0.01 and rect.y1 <= bottom_area + 0.6:
                    raise _Conflict("image crosses a removed band")
                begin = last_name[0] if last_name else start
                edits.append((begin, begin, wrap(ctm, fitz.Matrix(1, 0, 0, 1, 0, dy))))
                edits.append((end, end, b"\nQ"))
                changed += 1
        if bt_start is not None and op in (b"Tj", b"TJ", b"'", b'"'):
            y = baseline_y(tlm)
            dy = _shift(moves, y) if moves and top_area <= y <= bottom_area else 0.0
            bt_dys.add(round(dy, 3))
            if any(top <= y + dy <= bottom for top, bottom in bands):
                bt_hit = True
        if op == b"ET" and bt_start is not None:
            if len(bt_dys) > 1:
                raise _Conflict("text block spans a gap")
            dy = next(iter(bt_dys), 0.0)
            if bt_hit or dy:
                shift = fitz.Matrix(1, 0, 0, 1, 0, dy)
                edits.append((bt_start, bt_start, wrap(bt_ctm, shift * squeeze if bt_hit else shift)))
                edits.append((end, end, b"\nQ"))
                changed += 1
            bt_start = None
        operands.clear()

    if not edits:
        return 0
    pieces, last = [], 0
    for s0, e0, rep in sorted(edits, key=lambda e: (e[0], e[1])):
        if s0 < last:
            raise _Conflict("overlapping edits")
        pieces.append(data[last:s0])
        pieces.append(rep)
        last = e0
    pieces.append(data[last:])
    xref = doc.get_new_xref()
    doc.update_object(xref, "<<>>")
    doc.update_stream(xref, b"".join(pieces))
    page.set_contents(xref)
    for link in page.get_links():                          # keep clickable areas with their text
        r = fitz.Rect(link["from"])
        if moves and r.y0 >= top_area and r.y1 <= bottom_area + 0.5:
            dy0, dy1 = _shift(moves, r.y0), _shift(moves, r.y1)
            if dy0 or dy1:
                link["from"] = fitz.Rect(r.x0, r.y0 + dy0, r.x1, r.y1 + dy1)
                page.update_link(link)
    return changed


def _render_matches(before: fitz.Page, after: fitz.Page, moves: list[Move], area: tuple[float, float]) -> bool:
    """Every row of the original page appears unchanged at its new position, and nothing
    else appears. Rows next to a removed band (where borders end) are allowed to differ."""
    import numpy as np
    a, b = _render_gray(before).astype(np.int16), _render_gray(after).astype(np.int16)
    if a.shape != b.shape:
        return False
    z = 1 / GRID
    top_px, bot_px = int(area[0] * z), min(a.shape[0], int(area[1] * z))
    mapped = np.zeros(a.shape[0], bool)
    edges = set()
    for m in moves:
        for y in (m.y0, m.y1):
            r = int(round((y + _shift(moves, y + (0.01 if y == m.y1 else -0.01))) * z))
            edges.update(range(r - 3, r + 4))
    bad = 0
    for r in range(top_px, bot_px):
        y = (r + 0.5) / z
        if any(m.dy < 0 and m.y0 <= y < m.y1 for m in moves):
            continue                                          # this row was removed
        t = int(round(r + _shift(moves, y) * z))
        if not (top_px <= t < bot_px):
            if (a[r, 4:-4] < INK).sum() > 2:                  # content pushed off the body area
                return False
            continue                                          # empty row: nothing to keep
        mapped[t] = True
        if t in edges:
            continue
        diff = np.abs(a[r] - b[t])
        bad += int((diff > 60).sum())
    for t in range(top_px, bot_px):                           # newly opened space must be empty
        if not mapped[t] and t not in edges and (b[t] < INK).sum() > 2:
            return False
    return bad <= max(40, a.shape[1] // 10)


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


def _photo_pages(doc: fitz.Document) -> frozenset:
    """Photo pages keep their layout: captions are placed from the original photo positions."""
    pages = set()
    for index, page in enumerate(doc):
        if any(_clean(l).startswith("Photographs:") for l in page.get_text().splitlines()):
            pages.add(index)
    return frozenset(pages)


def apply_section_boxes(pdf_bytes: bytes, tighten: bool = False, boxes: bool = True,
                        protect_pages=()) -> BoxResult:
    """Box the summary subsections and, with tighten=True, close up extra space first:
    oversized gaps in the summary sections and empty space at the bottom of table cells
    (Appendix, Executive Summary, location plans). Photo pages are never moved."""
    skip: frozenset = frozenset()
    source = fitz.open(stream=pdf_bytes, filetype="pdf")
    protect = frozenset(protect_pages) | _photo_pages(source)
    try:
        for _attempt in range(len(source) + 1):
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            try:
                plan = plan_layout(doc, tighten=tighten, skip_pages=skip, boxed=boxes, protect_pages=protect)
                if not (plan.boxes and boxes) and not plan.moves:
                    return BoxResult(pdf_bytes, [], [], plan.notes, [])
                # 1. Prove each moved page renders the same, just closed up.
                failed = None
                for page_index in sorted({m.page_index for m in plan.moves}):
                    page_moves = [m for m in plan.moves if m.page_index == page_index]
                    area = _content_area(source[page_index])
                    scratch = fitz.open(stream=pdf_bytes, filetype="pdf")
                    try:
                        _rewrite_page(scratch[page_index], [], page_moves, area)
                        ok = _render_matches(source[page_index], scratch[page_index], page_moves, area)
                    except _Conflict:
                        ok = False
                    finally:
                        scratch.close()
                    if not ok:
                        failed = page_index
                        break
                if failed is not None:
                    skip = skip | {failed}
                    continue
                # 2. Apply moves, inset text and draw boxes.
                pages = {p.page_index for p in plan.boxes} if boxes else set()
                pages |= {m.page_index for m in plan.moves}
                for page_index in sorted(pages):
                    page = doc[page_index]
                    page_boxes = [p for p in plan.boxes if p.page_index == page_index] if boxes else []
                    bands = [(p.rect[1] - 0.5, p.rect[3] + 0.5) for p in page_boxes]
                    try:
                        _rewrite_page(page, bands, [m for m in plan.moves if m.page_index == page_index],
                                      _content_area(page))
                    except _Conflict:
                        failed = page_index
                        break
                    for box in page_boxes:
                        _draw_box(page, fitz.Rect(box.rect))
                if failed is not None:
                    skip = skip | {failed}
                    continue
                output = doc.tobytes(garbage=3, deflate=True)
            finally:
                doc.close()
            return BoxResult(output, plan.boxes if boxes else [], plan.titles if boxes else [],
                             plan.notes, plan.moves)
    finally:
        source.close()
    return BoxResult(pdf_bytes, [], [], ["Spacing and boxes were left unchanged (the layout couldn\u2019t be edited safely)."], [])
