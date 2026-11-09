"""PDF photo-caption detection and editing for TT site visit reports.

The engine intentionally edits only text associated with large images on pages
headed "Photographs:". All edits are returned as a new PDF byte string.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import io
import re
from typing import Iterable

import fitz

ZERO_WIDTH = "\u200b\u200c\u200d\ufeff"
ENGINE_VERSION = "2026-11-09.1"   # shown in the app so you can confirm what's deployed

# --- Caption font: Arial -----------------------------------------------------------
# Arial is licensed by Microsoft and isn't installed on Streamlit Cloud, so the engine
# looks for it in this order and reports which font was used:
#   1. assets/fonts/Arial.ttf (add it to the repo if TT's font licence allows)
#   2. Arial installed on the computer running the app (Windows / macOS / Linux)
#   3. the Arial embedded in the uploaded report (TT reports use ArialMT), if it has
#      every character the captions need and its letters map correctly
#   4. Liberation Sans (bundled, SIL OFL): an Arial look-alike with identical widths
FONT_DIR = Path(__file__).resolve().parent / "assets" / "fonts"
_ARIAL_FILES = (
    FONT_DIR / "Arial.ttf", FONT_DIR / "arial.ttf",
    Path("C:/Windows/Fonts/arial.ttf"),
    Path("/Library/Fonts/Arial.ttf"), Path("/System/Library/Fonts/Supplemental/Arial.ttf"),
    Path("/usr/share/fonts/truetype/msttcorefonts/Arial.ttf"),
    Path("/usr/share/fonts/truetype/msttcorefonts/arial.ttf"),
)
_ARIAL_BOLD_FILES = (
    FONT_DIR / "Arial Bold.ttf", FONT_DIR / "arialbd.ttf", FONT_DIR / "Arial-Bold.ttf",
    Path("C:/Windows/Fonts/arialbd.ttf"),
    Path("/Library/Fonts/Arial Bold.ttf"), Path("/System/Library/Fonts/Supplemental/Arial Bold.ttf"),
    Path("/usr/share/fonts/truetype/msttcorefonts/Arial_Bold.ttf"),
    Path("/usr/share/fonts/truetype/msttcorefonts/arialbd.ttf"),
)
_ARIAL_METRIC_FILE = FONT_DIR / "LiberationSans-Regular.ttf"
_ARIAL_BOLD_METRIC_FILE = FONT_DIR / "LiberationSans-Bold.ttf"
_ARIAL_NAME = re.compile(r"arial", re.I)
_NOT_REGULAR = re.compile(r"bold|italic|oblique|black|narrow|light", re.I)
_NOT_PLAIN_BOLD = re.compile(r"italic|oblique|black|narrow|light", re.I)
_BOLD = re.compile(r"bold", re.I)
_PDF_FONT_NAME = "TTArial"
_PDF_BOLD_FONT_NAME = "TTArialBd"


@dataclass
class CaptionFont:
    buffer: bytes | None   # None = PDF built-in Helvetica (last resort)
    label: str             # e.g. "Arial" or "Liberation Sans (Arial look-alike)"
    is_arial: bool
    bold_buffer: bytes | None = None   # used for the caption prefix; None = Helvetica-Bold

    @property
    def pdf_name(self) -> str:
        return _PDF_FONT_NAME if self.buffer else "helv"

    @property
    def pdf_bold_name(self) -> str:
        return _PDF_BOLD_FONT_NAME if self.bold_buffer else "hebo"

    def regular(self) -> fitz.Font:
        return fitz.Font(fontbuffer=self.buffer) if self.buffer else fitz.Font("helv")

    def bold(self) -> fitz.Font:
        return fitz.Font(fontbuffer=self.bold_buffer) if self.bold_buffer else fitz.Font("hebo")


def _font_covers(font: fitz.Font, chars: set[str]) -> bool:
    return all(font.has_glyph(ord(c)) for c in chars)


def _embedded_arial(doc: fitz.Document, chars: set[str], bold: bool = False) -> bytes | None:
    """Regular-weight Arial embedded in the report, if complete and correctly mapped.

    Embedded fonts are usually subsets, so every caption character must be present.
    Each character's width is also compared with Liberation Sans (same widths as
    Arial) so a subset whose characters are mapped to the wrong letters is rejected.
    """
    metric_file = _ARIAL_BOLD_METRIC_FILE if bold else _ARIAL_METRIC_FILE
    if not metric_file.exists():
        return None
    reference = fitz.Font(fontfile=str(metric_file))
    tried: set[int] = set()
    for page in doc:
        for xref, ext, _type, basefont, *_ in page.get_fonts(full=True):
            name = basefont.split("+")[-1]
            wrong_style = (
                (not _BOLD.search(name) or _NOT_PLAIN_BOLD.search(name)) if bold else _NOT_REGULAR.search(name)
            )
            if xref in tried or not _ARIAL_NAME.search(name) or wrong_style:
                continue
            tried.add(xref)
            try:
                _n, ext, _t, buffer = doc.extract_font(xref)
                if not buffer or ext in ("n/a", ""):
                    continue
                font = fitz.Font(fontbuffer=buffer)
            except Exception:  # noqa: BLE001 - unusable embedded font, try the next one
                continue
            if _font_covers(font, chars) and all(
                abs(font.glyph_advance(ord(c)) - reference.glyph_advance(ord(c))) < 0.005 for c in chars
            ):
                return buffer
    return None


def _font_file(paths: Iterable[Path], chars: set[str]) -> bytes | None:
    for path in paths:
        try:
            if path.exists():
                buffer = path.read_bytes()
                if _font_covers(fitz.Font(fontbuffer=buffer), chars):
                    return buffer
        except Exception:  # noqa: BLE001
            continue
    return None


def _resolve_style(doc: fitz.Document, chars: set[str], bold: bool) -> tuple[bytes | None, str, bool]:
    """(font bytes, label, is real Arial) for the regular or bold caption style."""
    buffer = _font_file(_ARIAL_BOLD_FILES if bold else _ARIAL_FILES, chars)
    if buffer:
        return buffer, "Arial", True
    embedded = _embedded_arial(doc, chars, bold=bold)
    if embedded:
        return embedded, "Arial (embedded in the report)", True
    metric = _ARIAL_BOLD_METRIC_FILE if bold else _ARIAL_METRIC_FILE
    if metric.exists():
        return metric.read_bytes(), "Liberation Sans (Arial look-alike)", False
    return None, "Helvetica", False


def resolve_caption_font(doc: fitz.Document, texts: Iterable[str]) -> CaptionFont:
    """Pick the caption fonts (Arial whenever possible): regular text plus a bold prefix."""
    texts = list(texts)
    chars = {c for text in texts for c in _split_caption(text)[1]}
    bold_chars = {c for text in texts for c in _split_caption(text)[0]}
    buffer, label, is_arial = _resolve_style(doc, chars, bold=False)
    bold_buffer, _bold_label, _bold_is_arial = _resolve_style(doc, bold_chars or set("0123456789.:-NCRGO"), bold=True)
    return CaptionFont(buffer, label, is_arial, bold_buffer)


def _split_caption(text: str) -> tuple[str, str]:
    """('65.NCR.1-1:', 'Description') - the prefix is set in bold."""
    match = re.match(r"^(\d+\.(?:NCR|GO)\.\d+-\d+:)\s*(.*)$", text, re.S)
    return (match.group(1), match.group(2)) if match else ("", text)


@dataclass
class CaptionJob:
    page_index: int
    photo_number: int
    observation_type: str | None
    observation_number: str
    caption: str
    image_rect: tuple[float, float, float, float]
    caption_rect: tuple[float, float, float, float]
    old_text_rects: list[tuple[float, float, float, float]] = field(default_factory=list)
    warning: str = ""

    @property
    def page_number(self) -> int:
        return self.page_index + 1


def clean_text(text: str) -> str:
    for char in ZERO_WIDTH:
        text = text.replace(char, "")
    return re.sub(r"\s+", " ", text).strip()


def extract_report_number(text: str) -> str:
    match = re.search(r"Special\s+Inspection\s+Report\s*#\s*(\d+)\b", clean_text(text), re.I)
    if match:
        return match.group(1)
    match = re.search(r"\bReport\s*(?:No\.?|#)\s*(\d+)\b", clean_text(text), re.I)
    return match.group(1) if match else ""


def parse_observation(text: str) -> tuple[str, str] | None:
    """Return (NCR|GO, observation number) only for detail-page headings."""
    text = clean_text(text)
    nc = re.search(
        r"Non[-\s]*Conformance\s+Observation\s*(?:#\s*)?(?:\d+\s*\.\s*NC\s*\.\s*)?(\d+)\b",
        text,
        re.I,
    )
    if nc:
        return "NCR", nc.group(1)
    go = re.search(
        r"General\s+Observation\s*(?:#\s*)?(?:(?:\d+\s*\.\s*GO\s*\.\s*)?)(\d+)\b",
        text,
        re.I,
    )
    if go:
        return "GO", go.group(1)
    return None


def format_caption(report_number: str, observation_type: str, observation_number: str,
                   photo_number: int, caption: str) -> str:
    """Format as REPORT.NCR/GO.OBS-n: existing caption (n is the photo number)."""
    kind = observation_type.upper().strip()
    if kind not in {"NCR", "GO"}:
        raise ValueError("Observation type must be NCR or GO")
    report = clean_text(report_number)
    obs = clean_text(observation_number)
    if not report or not re.fullmatch(r"\d+", report):
        raise ValueError("Report number must be numeric")
    if not obs or not re.fullmatch(r"\d+", obs):
        raise ValueError("Observation number must be numeric")
    photo_text = clean_text(caption)
    prefix = f"{report}.{kind}.{obs}-{photo_number}:"
    return f"{prefix} {photo_text}" if photo_text else prefix


def _photo_rectangles(page: fitz.Page) -> list[fitz.Rect]:
    """Return only likely photographic images; small logos/icons are ignored."""
    rects: list[fitz.Rect] = []
    seen: set[tuple[float, float, float, float]] = set()
    for image in page.get_images(full=True):
        xref = image[0]
        for rect in page.get_image_rects(xref):
            rect = fitz.Rect(rect)
            key = tuple(round(v, 1) for v in rect)
            if key in seen:
                continue
            seen.add(key)
            if rect.width >= 100 and rect.height >= 80 and rect.width * rect.height >= 15_000:
                rects.append(rect)
    return sorted(rects, key=lambda r: (round(r.y0, 1), r.x0))


def _text_lines(page: fitz.Page) -> list[tuple[fitz.Rect, str]]:
    lines: list[tuple[fitz.Rect, str]] = []
    for block in page.get_text("dict").get("blocks", []):
        for line in block.get("lines", []):
            text = clean_text("".join(span.get("text", "") for span in line.get("spans", [])))
            if text:
                spans = line.get("spans", [])
                rect = fitz.Rect(spans[0]["bbox"])
                for span in spans[1:]:
                    rect.include_rect(fitz.Rect(span["bbox"]))
                lines.append((rect, text))
    return lines


def _caption_lines_for_photo(
    lines: list[tuple[fitz.Rect, str]], image: fitz.Rect, next_row_top: float, footer_y: float | None,
) -> list[tuple[fitz.Rect, str]]:
    """Find caption lines directly under a photo, excluding page footers."""
    first_line_limit = image.y1 + 9.0
    final_limit = min(next_row_top - 1.5, image.y1 + 31.0)
    if footer_y is not None:
        final_limit = min(final_limit, footer_y - 1.5)
    candidates = []
    for rect, text in lines:
        if rect.y0 < image.y1 - 1.0 or rect.y0 > final_limit:
            continue
        if rect.x0 < image.x0 - 6.0 or rect.x0 > image.x0 + 16.0:
            continue
        if rect.x1 > image.x1 + 8.0:
            continue
        candidates.append((rect, text))
    first = next((item for item in candidates if item[0].y0 <= first_line_limit), None)
    if not first:
        return []
    # Keep only subsequent wrapped lines in the same short caption band;
    # do not absorb the report footer into an otherwise short caption.
    wrapped_limit = min(final_limit, first[0].y0 + 18.0)
    return [item for item in candidates if first[0].y0 - 1.0 <= item[0].y0 <= wrapped_limit]


def _rect_tuple(rect: fitz.Rect) -> tuple[float, float, float, float]:
    return tuple(float(v) for v in rect)


def analyze_pdf(pdf_bytes: bytes) -> tuple[str, list[CaptionJob], list[str]]:
    """Detect report number and caption jobs from an uploaded report PDF."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    report_number = ""
    jobs: list[CaptionJob] = []
    notes: list[str] = []
    current_observation: tuple[str, str] | None = None
    photo_counts: dict[tuple[str, str], int] = {}
    unclassified_counts: dict[int, int] = {}

    for page_index, page in enumerate(doc):
        text = page.get_text()
        if not report_number:
            report_number = extract_report_number(text)
        parsed = parse_observation(text)
        if parsed:
            current_observation = parsed
            photo_counts.setdefault(parsed, 0)

        if not re.search(r"\bPhotographs\s*:", clean_text(text), re.I):
            continue

        images = _photo_rectangles(page)
        if not images:
            notes.append(f"Page {page_index + 1}: found 'Photographs:' but no photo-sized images.")
            continue

        lines = _text_lines(page)
        footer_lines = [
            rect.y0 for rect, line_text in lines
            if rect.y0 > min(r.y1 for r in images)
            and re.search(r"Curtain Wall Inspection Report|Document Number", line_text, re.I)
        ]
        footer_y = min(footer_lines) if footer_lines else None
        for image in images:
            later_rows = [r.y0 for r in images if r.y0 > image.y0 + 5.0]
            next_row_top = min(later_rows) if later_rows else page.rect.height - 28.0
            caption_lines = _caption_lines_for_photo(lines, image, next_row_top, footer_y)
            caption = " ".join(text for _, text in caption_lines)
            old_rects = [_rect_tuple(rect) for rect, _ in caption_lines]

            if current_observation:
                key = current_observation
                photo_counts[key] = photo_counts.get(key, 0) + 1
                photo_number = photo_counts[key]
                kind, obs_number = key
                warning = ""
            else:
                unclassified_counts[page_index] = unclassified_counts.get(page_index, 0) + 1
                photo_number = unclassified_counts[page_index]
                kind, obs_number = None, ""
                warning = "Observation type/number not detected; review before export."

            # Reserve enough room before the next photo row or report footer.
            bottom = min(next_row_top - 1.5, page.rect.height - 33.0)
            if footer_y is not None and footer_y > image.y1:
                bottom = min(bottom, footer_y - 6.0)
            box = fitz.Rect(image.x0, image.y1 + 1.0, image.x1, bottom)
            if box.height < 7:
                warning = (warning + " Caption area is unusually short; review layout.").strip()
            jobs.append(CaptionJob(
                page_index=page_index,
                photo_number=photo_number,
                observation_type=kind,
                observation_number=obs_number,
                caption=caption,
                image_rect=_rect_tuple(image),
                caption_rect=_rect_tuple(box),
                old_text_rects=old_rects,
                warning=warning,
            ))

    if not report_number:
        notes.append("Could not detect a numeric Special Inspection Report number; enter it manually.")
    if not jobs:
        notes.append("No captionable photos detected. Pages must contain a 'Photographs:' heading and embedded photo images.")
    doc.close()
    return report_number, jobs, notes


def _existing_caption_without_prefix(caption: str) -> str:
    caption = clean_text(caption)
    return re.sub(r"^\d+\.(?:NCR|GO)\.\d+-(?:PHOTO\s+)?\d+\s*:\s*", "", caption, flags=re.I)


_LINE_HEIGHT = 1.0   # baseline-to-baseline distance, as a multiple of the font size


def _layout_caption(text: str, rect: fitz.Rect, size: float, font: CaptionFont | None,
                    ) -> list[tuple[float, float, str, bool]] | None:
    """Word-wrap a caption into `rect`: list of (x, baseline, text, bold), or None if it won't fit.

    The "65.NCR.1-1:" prefix is bold and the description regular, left-aligned.
    """
    font = font or CaptionFont(None, "Helvetica", False)
    regular, bold = font.regular(), font.bold()
    prefix, description = _split_caption(clean_text(text))
    words = ([(prefix, True)] if prefix else []) + [(w, False) for w in description.split()]
    if not words:
        return []
    space = regular.text_length(" ", fontsize=size)
    lines: list[list[tuple[str, bool, float]]] = [[]]
    line_width = 0.0
    for word, is_bold in words:
        width = (bold if is_bold else regular).text_length(word, fontsize=size)
        if width > rect.width + 0.01:
            return None   # a single word wider than the caption area
        needed = width if not lines[-1] else line_width + space + width
        if lines[-1] and needed > rect.width + 0.01:
            lines.append([])
            needed = width
        lines[-1].append((word, is_bold, needed - width))
        line_width = needed
    ascender = max(regular.ascender, bold.ascender)
    descender = min(regular.descender, bold.descender)
    first_baseline = rect.y0 + ascender * size
    last_bottom = first_baseline + (len(lines) - 1) * size * _LINE_HEIGHT - descender * size
    if last_bottom > rect.y1 + 0.01:
        return None
    placed = []
    for row, line in enumerate(lines):
        baseline = first_baseline + row * size * _LINE_HEIGHT
        # Merge neighbouring words of the same weight into one run.
        for word, is_bold, x in line:
            if placed and placed[-1][1] == baseline and placed[-1][3] == is_bold:
                px, pb, ptext, pbold = placed[-1]
                placed[-1] = (px, pb, f"{ptext} {word}", pbold)
            else:
                text_run = word if not placed or placed[-1][1] != baseline else f" {word}"
                if text_run.startswith(" "):
                    x -= space
                placed.append((rect.x0 + x, baseline, text_run, is_bold))
    return placed


def _fit_font_size(text: str, rect: fitz.Rect, max_size: float, min_size: float,
                   font: CaptionFont | None = None) -> float | None:
    """Largest size (0.25 pt steps) at which the caption fits its area, else None."""
    size = max_size
    while size >= min_size - 1e-6:
        if _layout_caption(text, rect, size, font) is not None:
            return size
        size = round(size - 0.25, 2)
    return None


# --- Page numbers -------------------------------------------------------------------
PAGE_NUMBER_SIZE = 8.0


@dataclass
class _PageNumber:
    text: str
    redact_rect: fitz.Rect
    right: float
    baseline: float
    size: float
    font: str
    color: tuple[float, float, float]


def _find_page_numbers(page: fitz.Page) -> list[_PageNumber]:
    """Every "Page X of Y" on a page, found character by character.

    Characters are grouped by baseline, so this works however the PDF splits the text
    into spans, lines or blocks (and in any PyMuPDF version).
    """
    rows: dict[tuple[float, float], list[tuple[str, fitz.Rect, dict]]] = {}
    for block in page.get_text("rawdict").get("blocks", []):
        for line in block.get("lines", []):
            if tuple(round(v, 2) for v in line.get("dir", (1, 0))) != (1.0, 0.0):
                continue   # horizontal text only
            for span in line.get("spans", []):
                for ch in span.get("chars", []):
                    key = (round(ch["origin"][1] * 2) / 2, round(span["size"], 1))
                    rows.setdefault(key, []).append((ch["c"], fitz.Rect(ch["bbox"]), span))
    found = []
    for (baseline, size), chars in rows.items():
        chars.sort(key=lambda item: item[1].x0)
        text, owners = "", []
        for i, (c, bbox, span) in enumerate(chars):
            if i and c != " " and text and text[-1] != " " and bbox.x0 - chars[i - 1][1].x1 > size * 0.15:
                text += " "; owners.append(None)   # visual gap without a space character
            text += c; owners.append(i)
        for match in re.finditer(r"Page\s*\d+\s*of\s*\d+", text, re.I):
            idx = [owners[k] for k in range(match.start(), match.end()) if owners[k] is not None]
            glyphs = [chars[k] for k in idx if chars[k][0].strip()]
            if not glyphs:
                continue
            span = glyphs[0][2]
            x0 = min(g[1].x0 for g in glyphs); x1 = max(g[1].x1 for g in glyphs)
            rect = fitz.Rect(x0 - 0.5, baseline - size * 0.8, x1 + 0.5, baseline + size * 0.25)
            rgb = span.get("color", 0)
            color = (((rgb >> 16) & 255) / 255, ((rgb >> 8) & 255) / 255, (rgb & 255) / 255)
            name = span.get("font", "").split("+")[-1]
            font = "hebo" if (span.get("flags", 0) & 16 or _BOLD.search(name)) else "helv"
            label = re.sub(r"\s+", " ", re.sub(r"(?i)page\s*(\d+)\s*of\s*(\d+)", r"Page \1 of \2", match.group(0)))
            found.append(_PageNumber(label, rect, x1, baseline, size, font, color))
    return found


def apply_captions(
    pdf_bytes: bytes,
    jobs: Iterable[CaptionJob],
    report_number: str,
    *,
    max_font_size: float = 10.5,
    min_font_size: float = 5.5,
    font: CaptionFont | None = None,
    page_number_size: float | None = PAGE_NUMBER_SIZE,
    stats: dict | None = None,
) -> tuple[bytes, list[str]]:
    """Apply captions to photo-only locations and return (PDF bytes, warnings).

    Captions are set in Arial when available (see resolve_caption_font); pass `font`
    to reuse a font already resolved with caption_font_for(). Every "Page X of Y"
    is rewritten at `page_number_size` (pass None to leave page numbers alone).
    If `stats` is a dict it receives page_numbers_found / page_numbers_resized / pages.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    jobs = list(jobs)
    if font is None:
        font = resolve_caption_font(doc, caption_texts(jobs, report_number))
    warnings: list[str] = []
    candidates = []
    cannot_fit: list[CaptionJob] = []

    for job in jobs:
        if job.observation_type not in {"NCR", "GO"} or not job.observation_number.strip():
            warnings.append(f"Page {job.page_number}, photo {job.photo_number}: skipped; observation is not classified.")
            continue
        if not job.caption.strip():
            warnings.append(f"Page {job.page_number}, photo {job.photo_number}: skipped; caption is blank.")
            continue
        try:
            caption_text = format_caption(
                report_number, job.observation_type, job.observation_number,
                job.photo_number, _existing_caption_without_prefix(job.caption),
            )
        except ValueError as exc:
            warnings.append(f"Page {job.page_number}, photo {job.photo_number}: {exc}")
            continue

        box = fitz.Rect(job.caption_rect)
        fitting_size = _fit_font_size(caption_text, box, max_font_size, min_font_size, font)
        if fitting_size is None:
            cannot_fit.append(job)
            continue
        candidates.append((job, caption_text, box, fitting_size))

    prepared = []
    if cannot_fit:
        first = cannot_fit[0]
        warnings.append(
            f"At least one caption cannot fit in its available area at {min_font_size:g} pt "
            f"(first: page {first.page_number}, photo {first.photo_number}); no generated captions "
            "were applied so the export does not contain a mix of caption sizes."
        )
    elif candidates:
        uniform_size = min(candidate[3] for candidate in candidates)
        prepared = [
            (job, caption_text, box, uniform_size)
            for job, caption_text, box, _ in candidates
        ]
        if uniform_size < max_font_size:
            warnings.append(
                f"Font size reduced to {uniform_size:g} pt for all {len(prepared)} generated photo captions "
                "to fit the tightest caption area."
            )

    pages_with_redactions: set[int] = set()
    for job, caption_text, box, font_size in prepared:
        page = doc[job.page_index]
        photo_bottom = fitz.Rect(job.image_rect).y1
        for old in job.old_text_rects:
            rect = fitz.Rect(old)
            rect.x0 = max(0, rect.x0 - 0.6)
            rect.y0 = max(0, rect.y0 - 0.5, photo_bottom)  # never paint over the photo itself
            rect.x1 = min(page.rect.width, rect.x1 + 0.6)
            rect.y1 = min(page.rect.height, rect.y1 + 0.5)
            # No fill: removing the old text is enough, and a white box would clip
            # the bottom edge of the photo when the old caption sits close to it.
            page.add_redact_annot(rect, fill=None, cross_out=False)
            pages_with_redactions.add(job.page_index)

    page_numbers: dict[int, list[_PageNumber]] = {}
    if page_number_size:
        for page_index, page in enumerate(doc):
            numbers = _find_page_numbers(page)
            if numbers:
                page_numbers[page_index] = numbers
                for number in numbers:
                    page.add_redact_annot(number.redact_rect, fill=None, cross_out=False)
                pages_with_redactions.add(page_index)

    # Remove only old caption / page-number text; leave images and vector artwork untouched.
    for page_index in pages_with_redactions:
        doc[page_index].apply_redactions(images=0, graphics=0, text=0)

    fonts_added: set[int] = set()
    for job, caption_text, box, font_size in prepared:
        page = doc[job.page_index]
        if job.page_index not in fonts_added:
            if font.buffer:
                page.insert_font(fontname=font.pdf_name, fontbuffer=font.buffer)
            if font.bold_buffer:
                page.insert_font(fontname=font.pdf_bold_name, fontbuffer=font.bold_buffer)
            fonts_added.add(job.page_index)
        runs = _layout_caption(caption_text, box, font_size, font)
        if runs is None:
            warnings.append(f"Page {job.page_number}, photo {job.photo_number}: unexpected text-fit issue during export.")
            continue
        for x, baseline, run_text, is_bold in runs:
            page.insert_text(
                (x, baseline), run_text, fontname=font.pdf_bold_name if is_bold else font.pdf_name,
                fontsize=font_size, color=(0.18, 0.18, 0.18), overlay=True,
            )

    for page_index, numbers in page_numbers.items():
        page = doc[page_index]
        for number in numbers:
            width = fitz.get_text_length(number.text, fontname=number.font, fontsize=page_number_size)
            page.insert_text(  # keep the right edge and baseline, so it stays bottom-right
                (number.right - width, number.baseline), number.text,
                fontname=number.font, fontsize=page_number_size, color=number.color, overlay=True,
            )

    if fonts_added:
        try:
            doc.subset_fonts()   # keep only the caption characters, so the PDF stays small
        except Exception:  # noqa: BLE001 - subsetting is a size optimisation only
            pass
    output = doc.tobytes(garbage=4, deflate=True)
    page_total = doc.page_count
    doc.close()

    found = sum(len(v) for v in page_numbers.values())
    resized = 0
    if page_number_size:
        # Check the saved PDF itself: each page number must now be at the new size, once.
        check = fitz.open(stream=output, filetype="pdf")
        for page_index, numbers in page_numbers.items():
            after = _find_page_numbers(check[page_index])
            for number in numbers:
                same = [a for a in after if a.text == number.text and abs(a.baseline - number.baseline) < 1.5]
                if same and all(abs(a.size - page_number_size) < 0.15 for a in same):
                    resized += 1
        check.close()
        if found and resized < found:
            warnings.append(f"Only {resized} of {found} page numbers could be resized to {page_number_size:g} pt.")
    if stats is not None:
        stats.update(page_numbers_found=found, page_numbers_resized=resized, pages=page_total)
    return output, warnings


def caption_texts(jobs: Iterable[CaptionJob], report_number: str) -> list[str]:
    """The caption strings that apply_captions would write (for font checks)."""
    texts = []
    for job in jobs:
        if job.observation_type in {"NCR", "GO"} and job.observation_number.strip() and job.caption.strip():
            try:
                texts.append(format_caption(report_number, job.observation_type, job.observation_number,
                                            job.photo_number, _existing_caption_without_prefix(job.caption)))
            except ValueError:
                pass
    return texts


def caption_font_for(pdf_bytes: bytes, jobs: Iterable[CaptionJob], report_number: str) -> CaptionFont:
    """Resolve the caption font for a report without editing it."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        return resolve_caption_font(doc, caption_texts(jobs, report_number))
    finally:
        doc.close()


# --- Signatures --------------------------------------------------------------------
MAX_SIGNATURES = 2
_SIGNED_LABEL = re.compile(r"^(Signed|Signatures?)\s*:", re.I)
_DATE_LABEL = re.compile(r"^Date\s*:", re.I)
_SIG_PAD = 2.0          # clearance from the cell border, in points
_SIG_LABEL_GAP = 6.0    # space between "Signed:" and the first signature
_SIG_GAP = 8.0          # space between two side-by-side signatures


@dataclass
class SignatureBox:
    page_index: int
    area: tuple[float, float, float, float]   # where signatures may go (inside the cell)
    cell: tuple[float, float, float, float]   # the whole "Signed:" cell
    date_cell: tuple[float, float, float, float] | None

    @property
    def page_number(self) -> int:
        return self.page_index + 1


def _filled_cells(page: fitz.Page) -> list[fitz.Rect]:
    cells = []
    for drawing in page.get_drawings():
        rect = fitz.Rect(drawing["rect"])
        if drawing.get("fill") is not None and rect.width > 20 and rect.height > 8:
            cells.append(rect)
    return cells


def find_signature_box(pdf_bytes: bytes) -> SignatureBox | None:
    """The space to the right of the report's "Signed:" label, inside its table cell."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        for page_index in range(doc.page_count - 1, -1, -1):   # the sign-off is near the end
            page = doc[page_index]
            lines = _text_lines(page)
            labels = [rect for rect, text in lines if _SIGNED_LABEL.match(text)]
            if not labels:
                continue
            label = labels[-1]
            # The "Signed:" text is a line by itself; measure the label word only.
            words = [fitz.Rect(w[:4]) for w in page.get_text("words")
                     if _SIGNED_LABEL.match(w[4] + (":" if not w[4].endswith(":") else ""))
                     and abs(w[1] - label.y0) < 3]
            label_rect = words[0] if words else label
            dates = [rect for rect, text in lines
                     if _DATE_LABEL.match(text) and abs(rect.y0 - label.y0) < 4 and rect.x0 > label_rect.x1]
            date_label = min(dates, key=lambda r: r.x0) if dates else None
            cells = [c for c in _filled_cells(page) if c.contains(label_rect)]
            cell = min(cells, key=lambda c: c.width * c.height) if cells else None
            date_cell = None
            if date_label is not None:
                date_cells = [c for c in _filled_cells(page) if c.contains(date_label)]
                date_cell = min(date_cells, key=lambda c: c.width * c.height) if date_cells else None
            if cell is None:   # no drawn cell: use the label's row up to the Date label
                right = (date_label.x0 - 6) if date_label is not None else label_rect.x0 + 200
                cell = fitz.Rect(label_rect.x0 - 1.5, label_rect.y0 - 4, right, label_rect.y1 + 4)
            right = cell.x1 - _SIG_PAD - 0.8                      # border lines are ~0.7 pt
            if date_cell is not None:
                right = min(right, date_cell.x0 - _SIG_PAD - 0.8)
            if date_label is not None:
                right = min(right, date_label.x0 - _SIG_PAD - 0.8)
            area = fitz.Rect(label_rect.x1 + _SIG_LABEL_GAP, cell.y0 + _SIG_PAD + 0.7,
                             right, cell.y1 - _SIG_PAD)
            if area.width < 20 or area.height < 6:
                continue
            return SignatureBox(page_index, _rect_tuple(area), _rect_tuple(cell),
                                _rect_tuple(date_cell) if date_cell is not None else None)
        return None
    finally:
        doc.close()


def prepare_signature(image_bytes: bytes) -> tuple[bytes, float]:
    """Trim blank margins and drop a white/off-white background. Returns (PNG, width/height)."""
    from PIL import Image, ImageOps

    image = ImageOps.exif_transpose(Image.open(io.BytesIO(image_bytes))).convert("RGBA")
    if max(image.size) > 2400:
        image.thumbnail((2400, 2400))
    pixels = image.load()
    width, height = image.size
    # Ink = clearly darker than the paper (or explicitly opaque on a transparent PNG).
    gray = ImageOps.grayscale(Image.alpha_composite(Image.new("RGBA", image.size, "white"), image))
    paper = sorted(gray.getdata())[int(width * height * 0.6)]          # typical background level
    threshold = min(paper - 40, 215)
    ink = gray.point(lambda v: 255 if v < threshold else 0)
    alpha = image.getchannel("A").point(lambda a: 255 if a > 40 else 0)
    from PIL import ImageChops
    mask = ImageChops.multiply(ink, alpha)
    box = mask.getbbox()
    if box is None:
        raise ValueError("no signature found in the image (it looks blank)")
    margin = max(2, int(0.02 * max(box[2] - box[0], box[3] - box[1])))
    box = (max(0, box[0] - margin), max(0, box[1] - margin),
           min(width, box[2] + margin), min(height, box[3] + margin))
    image = image.crop(box)
    gray = gray.crop(box)
    # Fade the paper to transparent; keep the ink's own colour.
    soft = gray.point(lambda v: 0 if v >= paper - 12 else (255 if v <= threshold else
                                int(255 * (paper - 12 - v) / max(1, paper - 12 - threshold))))
    image.putalpha(ImageChops.multiply(image.getchannel("A"), soft))
    if max(image.size) > 1200:
        image.thumbnail((1200, 1200))
    out = io.BytesIO()
    image.save(out, format="PNG", optimize=True)
    return out.getvalue(), image.width / image.height


def layout_signatures(area: fitz.Rect, ratios: list[float]) -> list[fitz.Rect]:
    """Scale signatures to fit `area`: one on its own, two side by side (left-aligned)."""
    area = fitz.Rect(area)
    if not ratios:
        return []
    slot_width = (area.width - _SIG_GAP * (len(ratios) - 1)) / len(ratios)
    rects = []
    x = area.x0
    for ratio in ratios:
        width = min(slot_width, area.height * ratio)
        height = width / ratio
        top = area.y0 + (area.height - height) / 2
        rects.append(fitz.Rect(x, top, x + width, top + height))
        x += slot_width + _SIG_GAP
    return rects


def apply_signatures(pdf_bytes: bytes, images: list[bytes]) -> tuple[bytes, list[str], dict]:
    """Place 1-2 signature images in the "Signed:" box. Returns (PDF, warnings, info)."""
    info = {"placed": 0, "page": None}
    if not images:
        return pdf_bytes, [], info
    if len(images) > MAX_SIGNATURES:
        return pdf_bytes, [f"Only {MAX_SIGNATURES} signatures fit; none were added."], info
    box = find_signature_box(pdf_bytes)
    if box is None:
        return pdf_bytes, ['No "Signed:" box was found in the report, so signatures were not added.'], info
    prepared, warnings = [], []
    for number, data in enumerate(images, start=1):
        try:
            prepared.append(prepare_signature(data))
        except Exception as exc:  # noqa: BLE001 - unreadable or blank image
            warnings.append(f"Signature {number} was not added: {exc}.")
    if not prepared:
        return pdf_bytes, warnings, info
    area = fitz.Rect(box.area)
    rects = layout_signatures(area, [ratio for _png, ratio in prepared])
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[box.page_index]
    for (png, _ratio), rect in zip(prepared, rects):
        page.insert_image(rect, stream=png, keep_proportion=True, overlay=True)
    output = doc.tobytes(garbage=3, deflate=True)
    doc.close()
    info.update(placed=len(prepared), page=box.page_number,
                rects=[_rect_tuple(r) for r in rects], area=box.area, date_cell=box.date_cell)
    return output, warnings, info



# --- Photo groups, previews and reordering ---------------------------------------
@dataclass
class PhotoGroup:
    """The photos of one observation ticket, in grid order (top-left, top-right, next row...)."""
    key: str
    observation_type: str | None          # NCR, GO or None when it wasn't detected
    observation_number: str
    jobs: list[int]                       # indexes into the job list, in slot order
    columns: int = 2                      # photos per grid row in the PDF
    _pages: tuple = ()

    @property
    def pages(self) -> list[int]:
        """Page numbers (1-based) the observation's photos are on."""
        return list(self._pages)

    def title(self, report_number: str = "") -> str:
        if self.observation_type in {"NCR", "GO"} and self.observation_number:
            ref = f"{report_number}.{self.observation_type}.{self.observation_number}" if report_number \
                else f"{self.observation_type} {self.observation_number}"
            kind = "Non-Conformance Observation" if self.observation_type == "NCR" else "General Observation"
            return f"{ref} \u00b7 {kind} {self.observation_number}"
        return "Photos not linked to an observation"


def group_photos(jobs: list[CaptionJob]) -> list[PhotoGroup]:
    """Group photos by the observation ticket they belong to, in report order.
    Photos with no detected observation are grouped by page."""
    groups: list[PhotoGroup] = []
    by_key: dict[str, PhotoGroup] = {}
    for index, job in enumerate(jobs):
        if job.observation_type in {"NCR", "GO"} and job.observation_number:
            key = f"{job.observation_type}.{job.observation_number}"
        else:
            key = f"page{job.page_index}"
        group = by_key.get(key)
        if group is None:
            group = PhotoGroup(key, job.observation_type, job.observation_number, [])
            by_key[key] = group
            groups.append(group)
        group.jobs.append(index)
    for group in groups:
        group._pages = tuple(sorted({jobs[i].page_index + 1 for i in group.jobs}))
        first_page = jobs[group.jobs[0]].page_index
        xs = {round(jobs[i].image_rect[0]) for i in group.jobs if jobs[i].page_index == first_page}
        all_xs = {round(j.image_rect[0]) for j in jobs if j.page_index == first_page}
        group.columns = max(1, min(3, len(all_xs) or len(xs)))
    return groups


def photo_previews(pdf_bytes: bytes, jobs: Iterable[CaptionJob], width: int = 240) -> list[bytes]:
    """Small JPEG previews of each photo (about `width` pixels wide)."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    previews = []
    try:
        for job in jobs:
            rect = fitz.Rect(job.image_rect)
            zoom = width / max(rect.width, 1)
            pix = doc[job.page_index].get_pixmap(clip=rect, matrix=fitz.Matrix(zoom, zoom), alpha=False)
            try:
                previews.append(pix.tobytes("jpg", jpg_quality=70))
            except Exception:  # noqa: BLE001 - older PyMuPDF without JPEG output
                previews.append(pix.tobytes("png"))
    finally:
        doc.close()
    return previews


def _image_draws(page: fitz.Page):
    """Each image drawn on the page: (name, byte range of 'name Do', CTM at the draw,
    page-space rect, graphics-state ops in effect). Also returns the stream and q depth."""
    from section_boxes import _tokens
    doc = page.parent
    data = b"\n".join(doc.xref_stream(x) or b"" for x in page.get_contents())
    to_page = page.transformation_matrix
    ctm, stack, gs = fitz.Matrix(1, 0, 0, 1, 0, 0), [], []
    draws, seg, last_name, depth = [], 0, None, 0
    for kind, value, start, end in _tokens(data):
        if kind != "op":
            if kind == "other" and data[start:start + 1] == b"/":
                last_name = (start, end)
            continue
        op = value
        operands = data[seg:start].split()
        if op == b"q":
            stack.append((ctm, list(gs)))
            depth += 1
        elif op == b"Q":
            if stack:
                ctm, gs = stack.pop()
            depth -= 1
        elif op == b"cm" and len(operands) >= 6:
            ctm = fitz.Matrix(*[float(v) for v in operands[-6:]]) * ctm
        elif op == b"gs":
            gs.append(data[seg:end].strip())
        elif op == b"Do" and last_name:
            m = ctm * to_page
            corners = [fitz.Point(x, y) * m for x in (0, 1) for y in (0, 1)]
            rect = fitz.Rect(min(p.x for p in corners), min(p.y for p in corners),
                             max(p.x for p in corners), max(p.y for p in corners))
            draws.append({"name": data[last_name[0] + 1:last_name[1]], "start": last_name[0], "end": end,
                          "ctm": ctm, "rect": rect, "gs": list(gs)})
        seg = end
    return data, draws, depth


def _xobject_is_image(doc: fitz.Document, page: fitz.Page, name: bytes) -> bool:
    import section_boxes as sb
    cat = sb._category(doc, sb._resources(doc, page), "XObject")
    if cat is None:
        return False
    kind, value = doc.xref_get_key(cat, name.decode("latin1"))
    if kind != "xref":
        return False
    return doc.xref_get_key(int(value.split()[0]), "Subtype")[1] == "/Image"


def reorder_photos(pdf_bytes: bytes, jobs: list[CaptionJob], orders: dict[str, list[int]],
                   groups: list[PhotoGroup]) -> tuple[bytes, list[CaptionJob], list[str]]:
    """Move photos to new grid positions within their observation.

    `orders[group.key]` lists the group's job indexes in the new order: the first goes in
    the first slot (top-left), the second in the next slot, and so on. A photo keeps its
    own caption text; its photo number becomes its new position. Returns the PDF, the
    jobs in their new slots (ready for apply_captions) and any notes."""
    import section_boxes as sb
    moves: list[tuple[int, int]] = []            # (job index of the photo, job index of its new slot)
    new_jobs = [replace_job(job) for job in jobs]
    for group in groups:
        order = orders.get(group.key) or group.jobs
        if sorted(order) != sorted(group.jobs):
            return pdf_bytes, jobs, [f"{group.title()}: the photo order wasn't valid, so it was left as is."]
        for slot, photo in zip(group.jobs, order):
            if slot != photo:
                moves.append((photo, slot))
            new_jobs[slot].caption = jobs[photo].caption
            new_jobs[slot].warning = jobs[photo].warning
    if not moves:
        return pdf_bytes, new_jobs, []

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        digests = {}
        pages = sorted({jobs[i].page_index for move in moves for i in move})
        for page_index in pages:
            for info in doc[page_index].get_image_info(hashes=True):
                digests[(page_index, tuple(round(v) for v in info["bbox"]))] = info["digest"]

        def digest(job: CaptionJob):
            return digests.get((job.page_index, tuple(round(v) for v in job.image_rect)))

        expected = {slot: digest(jobs[photo]) for photo, slot in moves}
        streams, found = {}, {}
        for page_index in pages:
            data, draws, depth = _image_draws(doc[page_index])
            streams[page_index] = [data, depth, [], []]          # data, depth, removals, placements
            for draw in draws:
                for index in {i for move in moves for i in move}:
                    job = jobs[index]
                    if job.page_index == page_index and abs(draw["rect"].x0 - job.image_rect[0]) < 1 \
                            and abs(draw["rect"].y0 - job.image_rect[1]) < 1 \
                            and abs(draw["rect"].x1 - job.image_rect[2]) < 1 \
                            and abs(draw["rect"].y1 - job.image_rect[3]) < 1:
                        if index in found:
                            raise ValueError("two images in one photo slot")
                        if not _xobject_is_image(doc, doc[page_index], draw["name"]):
                            raise ValueError("photo isn't a plain image")
                        found[index] = draw
        if len(found) != len({i for move in moves for i in move}):
            return pdf_bytes, jobs, ["Some photos couldn\u2019t be located in the PDF, so the photo order was left as is."]

        for photo, slot in moves:
            src, dst = jobs[photo], jobs[slot]
            draw = found[photo]
            old, cell = fitz.Rect(src.image_rect), fitz.Rect(dst.image_rect)
            # Same size as the slot's photo in PlanRadar's grid; otherwise fit inside the slot.
            scale = min(cell.width / old.width, cell.height / old.height, 1.0) \
                if (abs(old.width - cell.width) > 0.5 or abs(old.height - cell.height) > 0.5) else 1.0
            if old.width * scale > cell.width + 0.5 or old.height * scale > cell.height + 0.5:
                scale = min(cell.width / old.width, cell.height / old.height)
            new = fitz.Rect(cell.x0, cell.y0, cell.x0 + old.width * scale, cell.y0 + old.height * scale)
            to_page = doc[src.page_index].transformation_matrix
            move = (fitz.Matrix(1, 0, 0, 1, -old.x0, -old.y0) * fitz.Matrix(scale, 0, 0, scale, 0, 0)
                    * fitz.Matrix(1, 0, 0, 1, new.x0, new.y0))
            target_page = doc[dst.page_index]
            ctm = draw["ctm"] * to_page * move * ~target_page.transformation_matrix
            names = {draw["name"]} | {n for g in draw["gs"] for n in re.findall(rb"/([^\s/\[\]()<>{}%]+)", g)}
            renames = sb._share_resources(doc, doc[src.page_index], target_page, names) \
                if src.page_index != dst.page_index else {}
            ops = sb._rename(b" ".join(draw["gs"]) + b"\n/" + draw["name"] + b" Do", renames)
            streams[src.page_index][2].append((draw["start"], draw["end"]))
            streams[dst.page_index][3].append(b"q " + sb._fmt(ctm) + b"\n" + ops + b"\nQ")
            new_jobs[slot].image_rect = _rect_tuple(new)

        for page_index, (data, depth, removals, placements) in streams.items():
            pieces, pos = [], 0
            for start, end in sorted(removals):
                pieces.append(data[pos:start])
                pos = end
            pieces.append(data[pos:])
            body = b"".join(pieces)
            stream = b"q\n" + body + b"\n" + b"Q\n" * max(depth, 0) + b"Q\n" + b"\n".join(placements) + b"\n"
            xref = doc.get_new_xref()
            doc.update_object(xref, "<<>>")
            doc.update_stream(xref, stream)
            doc[page_index].set_contents(xref)
        output = doc.tobytes(garbage=3, deflate=True)
    except (ValueError, RuntimeError) as exc:
        return pdf_bytes, jobs, [f"The photo order couldn\u2019t be changed safely ({exc}), so it was left as is."]
    finally:
        doc.close()

    # Check: every moved slot now shows the photo that was put there, and nothing else changed.
    check = fitz.open(stream=output, filetype="pdf")
    try:
        for photo, slot in moves:
            job = new_jobs[slot]
            shown = [i["digest"] for i in check[job.page_index].get_image_info(hashes=True)
                     if fitz.Rect(i["bbox"]).contains(_middle(job.image_rect))]
            if shown != [expected[slot]]:
                return pdf_bytes, jobs, ["The photo order couldn\u2019t be verified after moving the photos, "
                                         "so it was left as is."]
    finally:
        check.close()
    moved = len(moves)
    return output, new_jobs, [f"Photo order changed: {moved} photo{'s' if moved != 1 else ''} moved to a new position."]


def _middle(rect) -> fitz.Point:
    r = fitz.Rect(rect)
    return fitz.Point((r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2)


def replace_job(job: CaptionJob) -> CaptionJob:
    from dataclasses import replace
    return replace(job, old_text_rects=list(job.old_text_rects))
