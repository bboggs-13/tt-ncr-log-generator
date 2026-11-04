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
    """Shift everything below a gap on one page by dy (negative = up)."""
    page_index: int
    after_y: float                 # bottom of the line above the gap
    before_baseline: float         # baseline of the first line below the gap
    dy: float
    old_gap: float
    new_gap: float
    label: str                     # first words of the line below the gap

    def note(self) -> str:
        verb = "tightened" if self.dy < 0 else "evened out"
        return (f"Page {self.page_index + 1}: {verb} the gap above "
                f"\u201c{self.label}\u201d from {self.old_gap:.1f} pt to {self.new_gap:.1f} pt.")


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
            if gap > limit:
                pass                                            # too big: tighten
            elif (a is not sec.head and b is not sec.stop and _ITEM_START.match(b.text)
                  and -1 < gap < item_gap - EVEN_UNDER):
                pass                                            # a list item with no gap: even out
            else:
                continue
            key = (b.page, round(b.baseline, 1))
            moves[key] = Move(b.page, a.rect.y1, b.baseline, round(target - gap, 2),
                              round(gap, 1), round(target, 1), _label(b.text))
    return sorted(moves.values(), key=lambda m: (m.page_index, m.before_baseline)), item_gap


def _dy(moves: list[Move], *, baseline: float | None = None, top: float | None = None) -> float:
    """Total shift for an element: text by its baseline, graphics by their top edge."""
    total = 0.0
    for m in moves:
        if (baseline is not None and baseline >= m.before_baseline - 0.5) or \
                (baseline is None and top is not None and top >= m.after_y + 0.5):
            total += m.dy
    return total


_UNMOVABLE = re.compile(rb"(?<![A-Za-z*'\"])(BI|sh)(?![A-Za-z*'\"])")


def _check_page_moves(page: fitz.Page, moves: list[Move], area: tuple[float, float],
                      lines: list[_Line]) -> tuple[list[Move], str | None]:
    """Drop moves that aren't safe: photos below the gap, or no room to space items out."""
    if not moves:
        return moves, None
    top_area, bottom_area = area
    first = min(m.after_y for m in moves)
    for info in page.get_image_info():
        r = fitz.Rect(info["bbox"])
        if r.y0 >= first and r.y0 >= top_area and r.y1 <= bottom_area + 1:
            return [], f"Page {page.number + 1}: spacing left as-is because a photo sits below the gap."
    stream = b"\n".join(page.parent.xref_stream(x) or b"" for x in page.get_contents())
    if _UNMOVABLE.search(stream):
        return [], f"Page {page.number + 1}: spacing left as-is (the page has content that can\u2019t be moved safely)."

    def overflows(candidate: list[Move]) -> bool:
        for line in lines:
            if not line.blank and line.rect.y1 + _dy(candidate, baseline=line.baseline) > bottom_area - ROOM_MARGIN:
                return True
        for drawing in page.get_drawings():
            r = drawing["rect"]
            if r.y0 >= top_area - 0.5 and r.y1 <= bottom_area + 0.5 and \
                    r.y1 + _dy(candidate, top=r.y0) > bottom_area + 0.5:
                return True
        return False

    if overflows(moves):
        kept = [m for m in moves if m.dy < 0]
        return kept, f"Page {page.number + 1}: not enough room to even out list spacing, so only large gaps were tightened."
    return moves, None


def plan_layout(doc: fitz.Document, tighten: bool = False, skip_pages: frozenset = frozenset(),
                boxed: bool = True) -> LayoutPlan:
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
    if tighten and sections:
        planned, _gap = _moves_for(sections, boxed)
        for page_index in sorted({m.page_index for m in planned}):
            page_moves = [m for m in planned if m.page_index == page_index]
            if page_index in skip_pages:
                notes.append(f"Page {page_index + 1}: spacing left as-is (the page\u2019s text couldn\u2019t be moved safely).")
                continue
            kept, note = _check_page_moves(doc[page_index], page_moves, areas[page_index],
                                           [l for l in flow if l.page == page_index])
            moves.extend(kept)
            if note:
                notes.append(note)
        # Shift the lines so the boxes are laid out where the text ends up.
        for line in flow:
            dy = _dy([m for m in moves if m.page_index == line.page], baseline=line.baseline)
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
    return LayoutPlan(plans, titles, moves, notes)


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
    """Inset text in the boxed bands (28.5-565.5 -> 30-564 pt) and move content below each gap.

    Text blocks are wrapped in their own transform; path coordinates are rewritten in place
    (so clipping paths still work). Raises _Conflict if something can't be moved whole.
    Returns the number of changed drawing units."""
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
    changed = 0

    def baseline_y(matrix: fitz.Matrix) -> float:
        return fitz.Point(0, 0).transform(matrix * bt_ctm * to_page).y

    def in_body(rect: fitz.Rect) -> bool:
        return rect.y0 >= top_area - 0.5 and rect.y1 <= bottom_area + 0.5

    for kind, value, start, end in _tokens(data):
        if kind == "num":
            operands.append((value, start, end))
            continue
        if kind == "other":
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
            m = path_ctm * to_page
            pts = []
            for pop, ops in path:
                nums = [o[0] for o in ops]
                if pop == b"re" and len(nums) >= 4:
                    x, y, w, h = nums[-4:]
                    pts += [fitz.Point(x, y), fitz.Point(x + w, y + h)]
                elif pop != b"h":
                    pts += [fitz.Point(nums[k], nums[k + 1]) for k in range(0, len(nums) - 1, 2)]
            if pts:
                pts = [p * m for p in pts]
                rect = fitz.Rect(min(p.x for p in pts), min(p.y for p in pts),
                                 max(p.x for p in pts), max(p.y for p in pts))
                dy = _dy(moves, top=rect.y0) if moves and in_body(rect) else 0.0
                if dy:
                    inv = ~m
                    d = fitz.Point(0, dy) * inv - fitz.Point(0, 0) * inv
                    for pop, ops in path:
                        coords = ops[-4:-2] if pop == b"re" else ops
                        if pop == b"re" and len(ops) < 4:
                            continue
                        for k, (v, s0, e0) in enumerate(coords):
                            edits.append((s0, e0, _num(v + (d.x if k % 2 == 0 else d.y))))
                    changed += 1
            path = []
        elif op == b"Do" and moves:
            corners = [fitz.Point(x, y) * (ctm * to_page) for x in (0, 1) for y in (0, 1)]
            rect = fitz.Rect(min(p.x for p in corners), min(p.y for p in corners),
                             max(p.x for p in corners), max(p.y for p in corners))
            if in_body(rect) and _dy(moves, top=rect.y0):
                raise _Conflict("image below a gap")
        if bt_start is not None and op in (b"Tj", b"TJ", b"'", b'"'):
            y = baseline_y(tlm)
            dy = _dy(moves, baseline=y) if moves and top_area <= y <= bottom_area else 0.0
            bt_dys.add(round(dy, 3))
            if any(top <= y + dy <= bottom for top, bottom in bands):
                bt_hit = True
        if op == b"ET" and bt_start is not None:
            if len(bt_dys) > 1:
                raise _Conflict("text block spans a gap")
            dy = next(iter(bt_dys), 0.0)
            if bt_hit or dy:
                shift = fitz.Matrix(1, 0, 0, 1, 0, dy)
                target = shift * squeeze if bt_hit else shift
                local = bt_ctm * (to_page * target * from_page) * ~bt_ctm
                edits.append((bt_start, bt_start, b"q " + _fmt(local) + b"\n"))
                edits.append((end, end, b"\nQ"))
                changed += 1
            bt_start = None
        operands.clear()

    if not edits:
        return 0
    pieces, last = [], 0
    for s0, e0, rep in sorted(edits, key=lambda e: (e[0], e[1])):
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
        dy = _dy(moves, top=r.y0) if moves and r.y1 <= bottom_area + 0.5 and r.y0 >= top_area else 0.0
        if dy:
            link["from"] = fitz.Rect(r.x0, r.y0 + dy, r.x1, r.y1 + dy)
            page.update_link(link)
    return changed


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


def apply_section_boxes(pdf_bytes: bytes, tighten: bool = False, boxes: bool = True) -> BoxResult:
    """Box the summary subsections and, with tighten=True, even out their spacing first."""
    skip: frozenset = frozenset()
    for _attempt in range(4):
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        try:
            plan = plan_layout(doc, tighten=tighten, skip_pages=skip, boxed=boxes)
            if not (plan.boxes and boxes) and not plan.moves:
                return BoxResult(pdf_bytes, [], [], plan.notes, [])
            conflict = None
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
                    conflict = page_index
                    break
                for box in page_boxes:
                    _draw_box(page, fitz.Rect(box.rect))
            if conflict is not None:
                skip = skip | {conflict}
                continue
            output = doc.tobytes(garbage=3, deflate=True)
        finally:
            doc.close()
        return BoxResult(output, plan.boxes if boxes else [], plan.titles if boxes else [],
                         plan.notes, plan.moves)
    return BoxResult(pdf_bytes, [], [], ["Spacing and boxes were left unchanged (the layout couldn\u2019t be edited safely)."], [])
