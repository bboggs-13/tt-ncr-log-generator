"""Non-conformance summary email from a TT special-inspection report PDF.

Reads the report text (no UI) and builds an email that pastes cleanly into Outlook:
  • subject line: project, report number, "Non-Conformance Summary"
  • one block per non-conformance: subject, status, responsible contractor, location,
    description, recommended action and required documentation
  • the previously reported NCRs listed as corrected in the Executive Summary

Where the details come from: each NCR's "Non-Conformance Observation <id>" ticket page only.
The end-of-report summary and Action Items lists are not read (reports no longer include a
non-conformance summary), so an NCR without a ticket page isn't in the email.

check_notes() lists things worth a second look before sending (shown in the app, never
in the email), e.g. an Executive Summary NCN number that doesn't match the report.
"""
from __future__ import annotations

import html
import re
from dataclasses import dataclass, field
from email.message import EmailMessage
from email.policy import SMTP

import fitz  # PyMuPDF

EMAIL_VERSION = "2026-11-17.1"

NCR_ID = r"\d+\.NCR?\.\d+"
_ID_RE = re.compile(NCR_ID)
_HEADING = re.compile(r"^[A-Z]\.\s+\S")                    # "C. Non-Conformance Observation ..."
_NOISE = re.compile(
    r"^(###issue\d+###|NOTICE: THIS SPECIAL INSPECTION REPORT|THAT IS PRIVILEGED, CONFIDENTIAL|"
    r"IMMEDIATELY BY EMAIL OR TELEPHONE|TT Project No:|Inspection Date:|Report Date:|"
    r"Special Inspection Report #\d+|Page \d+ of \d+|Curtain Wall Inspection Report|Document Number\s*[–-])")
_INVISIBLE = dict.fromkeys(map(ord, "\u200b\u200c\u200d\u2060\ufeff"), None)

DETAIL_LABELS = ["TT Inspector", "Contractor", "Status", "Subject", "Description",
                 "Recommended Action", "Required Documentation", "Location", "Floor",
                 "Elevation", "Grid Lines"]


@dataclass
class NcrItem:
    id: str
    subject: str = ""
    status: str = ""
    contractor: str = ""
    description: str = ""
    recommended_action: str = ""
    required_documentation: str = ""
    floor: str = ""
    elevation: str = ""
    grid_lines: str = ""
    sources: list = field(default_factory=list)   # "detail page"

    @property
    def location(self) -> str:
        parts = []
        floor = _blank_na(self.floor)
        if floor:
            parts.append(f"Floor {floor}" if re.fullmatch(r"\d+", floor) else floor)
        elev = _blank_na(self.elevation)
        if elev:
            parts.append(elev if "elevation" in elev.lower() else f"{elev} Elevation")
        grid = _blank_na(self.grid_lines)
        if grid:
            parts.append(f"Grid Lines {grid}")
        return ", ".join(parts)


@dataclass
class ReportSummary:
    project: str = ""
    report_no: str = ""
    inspection_type: str = ""        # e.g. "Special Inspection", from the "... Report #67" heading
    inspection_date: str = ""
    visit_date: str = ""             # from "On 10/30/2026 a representative ... visited" (cross-check only)
    report_date: str = ""
    tt_project_no: str = ""
    inspector: str = ""
    observed_flag: str = ""          # "Yes" / "No" from "Any Non-Conformances Observed"
    observed_ids: list = field(default_factory=list)
    corrected_flag: str = ""
    corrected_ids: list = field(default_factory=list)
    items: list = field(default_factory=list)


@dataclass
class Email:
    subject: str
    html: str
    text: str


# ----------------------------------------------------------------------------- reading
def _blank_na(value: str) -> str:
    v = (value or "").strip()
    return "" if v.upper() in ("", "N/A", "NA", "-", "--") else v


def _clean(text: str) -> str:
    text = text.translate(_INVISIBLE).replace("\u00a0", " ")
    return re.sub(r"\s+", " ", text).strip()


def _lines(pdf_bytes: bytes) -> tuple[list[str], str]:
    """Report body lines (headers/footers removed, pages joined) and the raw first page."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    out, first = [], doc[0].get_text() if len(doc) else ""
    for page in doc:
        for line in page.get_text().splitlines():
            line = line.translate(_INVISIBLE).replace("\u00a0", " ").strip()
            if line and not _NOISE.match(line):
                out.append(line)
    return out, first.translate(_INVISIBLE)


_ANY_LABEL = re.compile(r"^[A-Z][A-Za-z/'’.()\- ]{1,45}:")


def _labelled(block: list[str], labels: list[str], one_line: bool = False) -> dict:
    """Split 'Label: value' lines (values may run over several lines) into a dict.

    one_line: for two-column header blocks, a value never continues onto the next line.
    """
    pattern = re.compile(r"^(" + "|".join(re.escape(l) for l in labels) + r")\s*:\s*(.*)$")
    found, current = {}, None
    for line in block:
        m = pattern.match(line)
        if m:
            current = m.group(1)
            found.setdefault(current, m.group(2))
            if found[current] != m.group(2):     # label seen before: keep the first value
                current = None
        elif current and not one_line and not _ANY_LABEL.match(line):
            found[current] += " " + line
        else:
            current = None
    return {k: _clean(v) for k, v in found.items()}


def _section_after(lines: list[str], start: int, stops: tuple = ()) -> list[str]:
    """Lines after `start` up to the next lettered heading, photo page or stop phrase."""
    block = []
    for line in lines[start + 1:]:
        if _HEADING.match(line) or line.startswith("Photographs:") or line.startswith(stops):
            break
        block.append(line)
    return block


def _id_list(text: str) -> list[str]:
    return list(dict.fromkeys(_ID_RE.findall(text.replace(" ", ""))))


def _sort_key(ncr_id: str):
    nums = [int(n) for n in re.findall(r"\d+", ncr_id)]
    return nums


def extract_summary(pdf_bytes: bytes) -> ReportSummary:
    lines, first_page = _lines(pdf_bytes)
    s = ReportSummary()
    raw_all = "\n".join(lines)

    # Report details (from the page header, which _lines strips, so read page 1 raw)
    m = re.search(r"([A-Z][A-Za-z]*(?: [A-Z][A-Za-z]*)*? Inspection) Report #\s*\d+", first_page)
    if m:
        s.inspection_type = m.group(1).strip()
    for pattern, attr in ((r"Special Inspection Report #\s*(\d+)", "report_no"),
                          (r"Inspection Date:\s*([\d/.-]+)", "inspection_date"),
                          (r"Report Date:\s*([\d/.-]+)", "report_date"),
                          (r"TT Project No:\s*([A-Z0-9.\-]+)", "tt_project_no")):
        m = re.search(pattern, first_page)
        if m:
            setattr(s, attr, m.group(1).strip())
    m = re.search(r"\bOn\s*\u200b?\s*(\d{1,2}/\d{1,2}/\d{2,4})\s*\u200b?\s*a representative", first_page)
    if m:
        s.visit_date = m.group(1)
    head = _labelled(lines[:60], ["Project", "TT Inspector", "Client", "Project Location"], one_line=True)
    s.project = head.get("Project", "") or head.get("Project Location", "")
    s.inspector = head.get("TT Inspector", "")

    # Executive Summary: NCRs observed / corrected
    for i, line in enumerate(lines):
        m = re.match(r"Any Non-Conformances (Observed|Corrected)\s*:\s*(.*)$", line)
        if not m:
            continue
        flag = _clean(m.group(2))
        ids = []
        for nxt in lines[i + 1:i + 3]:
            n = re.match(r"NCN\s*#?\s*:\s*(.*)$", nxt)
            if n:
                ids = _id_list(n.group(1))
                break
        if m.group(1) == "Observed" and not s.observed_flag:
            s.observed_flag, s.observed_ids = flag, ids
        elif m.group(1) == "Corrected" and not s.corrected_flag:
            s.corrected_flag, s.corrected_ids = flag, ids

    items: dict[str, NcrItem] = {}

    def item(ncr_id: str) -> NcrItem:
        return items.setdefault(ncr_id, NcrItem(id=ncr_id))

    # Non-Conformance Observation ticket pages
    for i, line in enumerate(lines):
        m = re.match(r"^[A-Z]\.\s+Non-Conformance Observation\s+(" + NCR_ID + r")", line.replace(" .", "."))
        if not m:
            continue
        fields = _labelled(_section_after(lines, i), DETAIL_LABELS)
        it = item(m.group(1))
        for attr, label in (("subject", "Subject"), ("status", "Status"), ("contractor", "Contractor"),
                            ("description", "Description"), ("recommended_action", "Recommended Action"),
                            ("required_documentation", "Required Documentation"), ("floor", "Floor"),
                            ("elevation", "Elevation"), ("grid_lines", "Grid Lines")):
            if fields.get(label) and not getattr(it, attr):
                setattr(it, attr, fields[label])
        it.sources.append("detail page")

    # Only the ticket pages are used; summary / Action Items lists are not read.
    s.items = sorted(items.values(), key=lambda it: _sort_key(it.id))
    return s


# ----------------------------------------------------------------------------- checks
def check_notes(s: ReportSummary) -> list[str]:
    """Things to double-check in the report before sending (never put in the email)."""
    notes = []
    found = [it.id for it in s.items]
    if s.observed_ids and set(s.observed_ids) != set(found):
        msg = f"The Executive Summary lists {', '.join(s.observed_ids)} as observed"
        if found:
            msg += f", but the report's non-conformances are {', '.join(found)}"
        notes.append(msg + ". The email uses the non-conformance ticket pages in the report.")
    if s.inspection_date and s.visit_date and s.inspection_date != s.visit_date:
        notes.append(f"The report's Inspection Date is {s.inspection_date}, but the description says the "
                     f"site was visited on {s.visit_date}. The opening sentence uses the Inspection Date; "
                     "edit it if the visit date is the right one.")
    if s.observed_flag.lower().startswith("no") and found:
        notes.append("The Executive Summary says no non-conformances were observed, "
                     f"but the report has {len(found)} non-conformance ticket page"
                     f"{'s' if len(found) != 1 else ''}.")
    if s.report_no:
        stray = [it.id for it in s.items if it.id.split(".")[0] != s.report_no]
        if stray:
            notes.append(f"{', '.join(stray)} {'does' if len(stray) == 1 else 'do'} not start with this "
                         f"report's number ({s.report_no}).")
    by_action = {}
    for it in s.items:
        if it.recommended_action:
            by_action.setdefault(it.recommended_action, []).append(it.id)
    for ids in by_action.values():
        if len(ids) > 1:
            notes.append(f"{', '.join(ids)} have the same Recommended Action word for word. "
                         "Check it wasn't copied over by mistake.")
    for it in s.items:
        if not it.description:
            notes.append(f"{it.id} has no description on its ticket page.")
        if not it.recommended_action:
            notes.append(f"{it.id} has no recommended action on its ticket page.")
    return notes


# ----------------------------------------------------------------------------- email
@dataclass
class EmailWording:
    greeting: str = "Hello all,"
    intro: str = ""                  # blank = default sentence built from the report
    closing: str = ("Please review the recommended actions above and submit the required "
                    "documentation for TT's review. Let us know if you have any questions or "
                    "would like to coordinate a follow-up review on site.")
    sign_off: str = "Thank you,"
    include_corrected: bool = True


def _with_weekday(date_text: str) -> str:
    """"11/10/2026" -> "Tuesday, 11/10/2026". Left as written if it isn't a recognisable date."""
    from datetime import datetime
    text = (date_text or "").strip()
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%m-%d-%Y", "%m.%d.%Y", "%Y-%m-%d"):
        try:
            return f"{datetime.strptime(text, fmt):%A}, {text}"
        except ValueError:
            continue
    return text


def default_intro(s: ReportSummary) -> str:
    report = f"Special Inspection Report #{s.report_no}" if s.report_no else "the attached special inspection report"
    when = f" on {s.inspection_date}" if s.inspection_date else ""
    if not s.items:
        return (f"No non-conformances were observed during TT's special inspection{when}, "
                f"as documented in {report}. The full report is attached for reference.")
    kind = s.inspection_type or "Special Inspection"
    report = f"the attached {kind} Report #{s.report_no}" if s.report_no else f"the attached {kind.lower()} report"
    visit = f" on {_with_weekday(s.inspection_date)}" if s.inspection_date else ""
    return (f"Please find {report} and associated NCR log for TT's visit{visit}. Documented below is a "
            f"summary of the non-conformances observed during TT's {kind.lower()}.")


def subject_line(s: ReportSummary) -> str:
    parts = [p for p in (s.project, f"Special Inspection Report #{s.report_no}" if s.report_no else "",
                         "Non-Conformance Summary") if p]
    return " \u2013 ".join(parts)


_FONT = "font-family:Aptos,Calibri,Arial,sans-serif;font-size:11pt;color:#000000;"
_P = f'<p style="margin:0 0 10pt 0;{_FONT}">'


def build_email(s: ReportSummary, wording: EmailWording | None = None) -> Email:
    w = wording or EmailWording()
    intro = w.intro.strip() or default_intro(s)
    e = html.escape
    html_parts = [f'<div style="{_FONT}">']
    text = []
    if w.greeting.strip():
        html_parts.append(f"{_P}{e(w.greeting.strip())}</p>")
        text += [w.greeting.strip(), ""]
    html_parts.append(f"{_P}{e(intro)}</p>")
    text += [intro, ""]

    for it in s.items:
        title = it.id + (f" \u2013 {it.subject}" if it.subject else "")
        meta = [(k, v) for k, v in (("Status", it.status), ("Responsible", it.contractor),
                                    ("Location", it.location)) if _blank_na(v)]
        rows = [(k, v) for k, v in (("Description", it.description),
                                    ("Recommended Action", it.recommended_action),
                                    ("Required Documentation", it.required_documentation)) if _blank_na(v)]
        html_parts.append(f'<p style="margin:12pt 0 4pt 0;{_FONT}"><b>{e(title)}</b></p>')
        text.append(title)
        if meta:
            html_parts.append(f'<p style="margin:0 0 4pt 0;{_FONT}">' +
                              " &nbsp;|&nbsp; ".join(f"<b>{e(k)}:</b> {e(v)}" for k, v in meta) + "</p>")
            text.append(" | ".join(f"{k}: {v}" for k, v in meta))
        for k, v in rows:
            html_parts.append(f'<p style="margin:0 0 4pt 0;{_FONT}"><b>{e(k)}:</b> {e(v)}</p>')
            text.append(f"{k}: {v}")
        text.append("")

    if w.include_corrected and s.corrected_ids:
        line = ("The following previously reported non-conformances were verified as corrected: "
                + ", ".join(s.corrected_ids) + ".")
        html_parts.append(f'<p style="margin:12pt 0 10pt 0;{_FONT}">{e(line)}</p>')
        text += [line, ""]

    if w.closing.strip():
        html_parts.append(f'<p style="margin:12pt 0 10pt 0;{_FONT}">{e(w.closing.strip())}</p>')
        text += [w.closing.strip(), ""]
    if w.sign_off.strip():
        html_parts.append(f"{_P}{e(w.sign_off.strip())}</p>")
        text.append(w.sign_off.strip())
    html_parts.append("</div>")
    return Email(subject=subject_line(s), html="".join(html_parts), text="\n".join(text).strip() + "\n")


def build_eml(email: Email) -> bytes:
    """Unsent .eml draft: opens in Outlook (desktop) as a new message ready to address."""
    msg = EmailMessage(policy=SMTP)
    msg["Subject"] = email.subject
    msg["X-Unsent"] = "1"
    msg.set_content(email.text)
    msg.add_alternative(f"<html><body>{email.html}</body></html>", subtype="html")
    return msg.as_bytes()
