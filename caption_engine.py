"""PDF photo-caption detection and editing for TT site visit reports.

The engine intentionally edits only text associated with large images on pages
headed "Photographs:". All edits are returned as a new PDF byte string.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import re
from typing import Iterable

import fitz

ZERO_WIDTH = "\u200b\u200c\u200d\ufeff"

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
_ARIAL_METRIC_FILE = FONT_DIR / "LiberationSans-Regular.ttf"
_ARIAL_NAME = re.compile(r"arial", re.I)
_NOT_REGULAR = re.compile(r"bold|italic|oblique|black|narrow|light", re.I)
_PDF_FONT_NAME = "TTArial"


@dataclass
class CaptionFont:
    buffer: bytes | None   # None = PDF built-in Helvetica (last resort)
    label: str             # e.g. "Arial" or "Liberation Sans (Arial look-alike)"
    is_arial: bool

    @property
    def pdf_name(self) -> str:
        return _PDF_FONT_NAME if self.buffer else "helv"


def _font_covers(font: fitz.Font, chars: set[str]) -> bool:
    return all(font.has_glyph(ord(c)) for c in chars)


def _embedded_arial(doc: fitz.Document, chars: set[str]) -> bytes | None:
    """Regular-weight Arial embedded in the report, if complete and correctly mapped.

    Embedded fonts are usually subsets, so every caption character must be present.
    Each character's width is also compared with Liberation Sans (same widths as
    Arial) so a subset whose characters are mapped to the wrong letters is rejected.
    """
    if not _ARIAL_METRIC_FILE.exists():
        return None
    reference = fitz.Font(fontfile=str(_ARIAL_METRIC_FILE))
    tried: set[int] = set()
    for page in doc:
        for xref, ext, _type, basefont, *_ in page.get_fonts(full=True):
            name = basefont.split("+")[-1]
            if xref in tried or not _ARIAL_NAME.search(name) or _NOT_REGULAR.search(name):
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


def resolve_caption_font(doc: fitz.Document, texts: Iterable[str]) -> CaptionFont:
    """Pick the caption font (Arial whenever possible) for this report's captions."""
    chars = {c for text in texts for c in text}
    for path in _ARIAL_FILES:
        try:
            if path.exists():
                buffer = path.read_bytes()
                if _font_covers(fitz.Font(fontbuffer=buffer), chars):
                    return CaptionFont(buffer, "Arial", True)
        except Exception:  # noqa: BLE001
            continue
    embedded = _embedded_arial(doc, chars)
    if embedded:
        return CaptionFont(embedded, "Arial (embedded in the report)", True)
    if _ARIAL_METRIC_FILE.exists():
        return CaptionFont(_ARIAL_METRIC_FILE.read_bytes(), "Liberation Sans (Arial look-alike)", False)
    return CaptionFont(None, "Helvetica", False)


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


def _fit_font_size(text: str, rect: fitz.Rect, max_size: float, min_size: float,
                   font: CaptionFont | None = None) -> float | None:
    """Find a fitting font size without drawing on the source page."""
    test_doc = fitz.open()
    test_page = test_doc.new_page(width=max(rect.width, 1), height=max(rect.height, 1))
    fontname = "helv"
    if font is not None and font.buffer:
        test_page.insert_font(fontname=font.pdf_name, fontbuffer=font.buffer)
        fontname = font.pdf_name
    local_rect = fitz.Rect(0, 0, rect.width, rect.height)
    size = max_size
    while size >= min_size - 1e-6:
        remaining = test_page.insert_textbox(
            local_rect, text, fontname=fontname, fontsize=size,
            lineheight=1.0, align=0, overlay=True,
        )
        if remaining >= -0.01:
            test_doc.close()
            return size
        size = round(size - 0.25, 2)
    test_doc.close()
    return None


def apply_captions(
    pdf_bytes: bytes,
    jobs: Iterable[CaptionJob],
    report_number: str,
    *,
    max_font_size: float = 10.5,
    min_font_size: float = 5.5,
    font: CaptionFont | None = None,
) -> tuple[bytes, list[str]]:
    """Apply captions to photo-only locations and return (PDF bytes, warnings).

    Captions are set in Arial when available (see resolve_caption_font); pass `font`
    to reuse a font already resolved with caption_font_for().
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

    # Remove only old caption text; leave images and vector artwork untouched.
    for page_index in pages_with_redactions:
        doc[page_index].apply_redactions(images=0, graphics=0, text=0)

    fonts_added: set[int] = set()
    for job, caption_text, box, font_size in prepared:
        page = doc[job.page_index]
        if font.buffer and job.page_index not in fonts_added:
            page.insert_font(fontname=font.pdf_name, fontbuffer=font.buffer)
            fonts_added.add(job.page_index)
        remaining = page.insert_textbox(
            box, caption_text, fontname=font.pdf_name, fontsize=font_size,
            lineheight=1.0, align=0, color=(0.18, 0.18, 0.18), overlay=True,
        )
        if remaining < -0.01:
            warnings.append(f"Page {job.page_number}, photo {job.photo_number}: unexpected text-fit issue during export.")

    if fonts_added:
        try:
            doc.subset_fonts()   # keep only the caption characters, so the PDF stays small
        except Exception:  # noqa: BLE001 - subsetting is a size optimisation only
            pass
    output = doc.tobytes(garbage=4, deflate=True)
    doc.close()
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
