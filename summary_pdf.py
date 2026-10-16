"""Time Sensitive NCR Summary — one-page PDF in the TT US Letter primary template style.

Layout mirrors PrimaryTemplates_USLetter_12ptFont_2023_rev: 0.75 in margins, 7 pt uppercase
running header with bold page number, 0.5 pt rule at 49 pt, TT logo bottom-right (126 pt wide).
Counts use the exact rules behind the Time Sensitive log highlighting (engine.AGE_BANDS):
open items only (anything not Closed), aged from Observation Date to the distribution date.
"""
from __future__ import annotations

import io
from datetime import date

import pandas as pd
from reportlab.lib.colors import CMYKColor, HexColor, white
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

import engine

PAGE_W, PAGE_H = letter                     # 612 x 792 pt
MARGIN = 54                                 # 0.75 in, per template
CONTENT_W = PAGE_W - 2 * MARGIN             # 504 pt
TT_GRAY = CMYKColor(0.40, 0.30, 0.20, 0.66)  # template header/rule colour
TT_OLIVE = HexColor("#839058")
TT_TERRA = HexColor("#D53822")
LIGHT_RULE = HexColor("#C9C9CC")
MUTED = HexColor("#6E6E76")
CLOSED_WINDOWS = (30, 60, 90)                # "closed in the last N days" (cumulative)
FONT, FONT_LIGHT, FONT_BOLD = "Helvetica", "Helvetica", "Helvetica-Bold"  # Univers stand-in


def _hex(fill) -> str:
    return "#" + fill.fgColor.rgb[-6:]


def _categories():
    """Four categories, matching the log legend and highlight colours."""
    cats = [{"label": "Less than 30 days", "lo": None, "hi": 30, "colour": None, "swatch": "No highlight"}]
    names = ["Yellow", "Orange", "Red"]
    for (lo, hi, _l, fill), name in zip(engine.AGE_BANDS, names):
        cats.append({"label": f"{lo}–{hi - 1} days" if hi else f"{lo}+ days",
                     "lo": lo, "hi": hi, "colour": _hex(fill), "swatch": name})
    return cats


def age_summary(df: pd.DataFrame, ref_date: date | None = None) -> dict:
    """Count open items per age band exactly as the Time Sensitive log highlights them."""
    ref = ref_date or date.today()
    obs_col = engine._resolve_field("Observation Date", df)
    st_col = engine._resolve_field("Status", df)
    status = df[st_col].astype(str).str.strip().str.lower() if st_col else pd.Series("", index=df.index)
    is_open = ~status.isin(engine.CLOSED_WORDS)
    obs = pd.to_datetime(df[obs_col], errors="coerce") if obs_col else pd.Series(pd.NaT, index=df.index)
    ages = (pd.Timestamp(ref) - obs).dt.days

    cats = _categories()
    open_ages = ages[is_open]
    dated = open_ages.dropna()
    for c in cats:
        m = pd.Series(True, index=dated.index)
        if c["lo"] is not None:
            m &= dated >= c["lo"]
        if c["hi"] is not None:
            m &= dated < c["hi"]
        c["count"] = int(m.sum())
    # Recently closed: Status = Closed with a valid Date Closed, within N days of distribution
    dc_col = engine._resolve_field("Date Closed", df)
    closed_dt = pd.to_datetime(df[dc_col], errors="coerce") if dc_col else pd.Series(pd.NaT, index=df.index)
    closed_ages = (pd.Timestamp(ref) - closed_dt).dt.days[~is_open]
    recent = {w: int(((closed_ages >= 0) & (closed_ages <= w)).sum()) for w in CLOSED_WINDOWS}
    closed_undated = int(closed_ages.isna().sum())

    oldest = None
    if len(dated):
        i = dated.idxmax()
        oldest = {"days": int(dated[i]), "date": obs[i].date()}
    return {"ref_date": ref, "categories": cats, "open_total": int(is_open.sum()),
            "closed": int((~is_open).sum()), "undated": int(open_ages.isna().sum()),
            "total": len(df), "oldest": oldest,
            "closed_recent": recent, "closed_undated": closed_undated}


# ---------------------------------------------------------------- drawing helpers
def _y(top: float) -> float:
    """Convert a top-down coordinate (like the template measurements) to PDF y."""
    return PAGE_H - top


def _section(c, top, title):
    c.setFillColor(TT_GRAY)
    c.setFont(FONT_BOLD, 9)
    c.drawString(MARGIN, _y(top), title.upper())
    c.setStrokeColor(TT_GRAY)
    c.setLineWidth(0.5)
    c.line(MARGIN, _y(top + 6), MARGIN + CONTENT_W, _y(top + 6))


def _fmt_date(d: date) -> str:
    return f"{d:%B} {d.day}, {d.year}"


def build_summary_pdf(df: pd.DataFrame, profile, project_name: str,
                      ref_date: date | None = None) -> bytes:
    s = age_summary(df, ref_date)
    cats, n_open = s["categories"], s["open_total"]
    kind = "Special Inspections" if profile.key.startswith("special") else "Regular Site Inspections"
    log_name = "NCR Log" if profile.key.startswith("special") else "Site Inspection Log"
    project = project_name.strip() or "Project"

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    c.setTitle(f"{project} — Time Sensitive {log_name} Summary")
    c.setAuthor("Thornton Tomasetti")

    # Running header + rule (template positions)
    c.setFillColor(TT_GRAY)
    c.setFont(FONT_LIGHT, 7)
    c.drawString(MARGIN, _y(41.2), f"{project.upper()}  |  TIME SENSITIVE {log_name.upper()} SUMMARY")
    c.setFont(FONT_BOLD, 7)
    c.drawRightString(MARGIN + CONTENT_W, _y(41.2), "1")
    c.setStrokeColor(TT_GRAY)
    c.setLineWidth(0.5)
    c.line(MARGIN, _y(49), MARGIN + CONTENT_W, _y(49))

    # Title block
    c.setFillColor(TT_GRAY)
    c.setFont(FONT_BOLD, 22)
    c.drawString(MARGIN, _y(100), "Time Sensitive NCR Summary")
    c.setFillColor(TT_OLIVE)
    c.setFont(FONT, 12)
    c.drawString(MARGIN, _y(119), f"{kind}  ·  Open items by days since observation")

    # Project / distribution details
    top = 150
    cols = [("Project", project), ("Date of distribution", _fmt_date(s["ref_date"])),
            ("Open items", str(n_open)), ("Closed items", str(s["closed"]))]
    widths = [210, 150, 72, 72]
    x = MARGIN
    for (label, value), w in zip(cols, widths):
        c.setFillColor(MUTED)
        c.setFont(FONT_BOLD, 7.5)
        c.drawString(x, _y(top), label.upper())
        c.setFillColor(TT_GRAY)
        c.setFont(FONT, 12)
        c.drawString(x, _y(top + 16), value)
        x += w
    c.setStrokeColor(LIGHT_RULE)
    c.setLineWidth(0.5)
    c.line(MARGIN, _y(top + 28), MARGIN + CONTENT_W, _y(top + 28))

    # Four category cards
    _section(c, 214, "Open items by age")
    gap = 12
    card_w = (CONTENT_W - 3 * gap) / 4
    card_top, card_h, band_h = 232, 118, 10
    for i, cat in enumerate(cats):
        x = MARGIN + i * (card_w + gap)
        c.setStrokeColor(LIGHT_RULE)
        c.setLineWidth(0.75)
        c.setFillColor(white)
        c.rect(x, _y(card_top + card_h), card_w, card_h, stroke=1, fill=1)
        if cat["colour"]:
            c.setFillColor(HexColor(cat["colour"]))
            c.rect(x, _y(card_top + band_h), card_w, band_h, stroke=0, fill=1)
        else:
            c.setStrokeColor(LIGHT_RULE)
            c.line(x, _y(card_top + band_h), x + card_w, _y(card_top + band_h))
        c.setFillColor(TT_GRAY)
        c.setFont(FONT_BOLD, 34)
        c.drawCentredString(x + card_w / 2, _y(card_top + 62), str(cat["count"]))
        c.setFont(FONT_BOLD, 10)
        c.drawCentredString(x + card_w / 2, _y(card_top + 84), cat["label"])
        pct = f"{cat['count'] / n_open:.0%} of open items" if n_open else "—"
        c.setFillColor(MUTED)
        c.setFont(FONT, 8)
        c.drawCentredString(x + card_w / 2, _y(card_top + 99), pct)
        c.drawCentredString(x + card_w / 2, _y(card_top + 110),
                            "No highlight" if not cat["colour"] else f"{cat['swatch']} in log")

    # Distribution bar
    bar_top, bar_h = 400, 16
    _section(c, 384, "Distribution of open items")
    c.setStrokeColor(LIGHT_RULE)
    c.setLineWidth(0.75)
    dated_total = sum(cat["count"] for cat in cats)
    x = MARGIN
    if dated_total:
        for cat in cats:
            w = CONTENT_W * cat["count"] / dated_total
            if w <= 0:
                continue
            c.setFillColor(HexColor(cat["colour"]) if cat["colour"] else white)
            c.rect(x, _y(bar_top + bar_h), w, bar_h, stroke=1, fill=1)
            if w > 22:
                c.setFillColor(TT_GRAY)
                c.setFont(FONT_BOLD, 8)
                c.drawCentredString(x + w / 2, _y(bar_top + 11), str(cat["count"]))
            x += w
    else:
        c.rect(MARGIN, _y(bar_top + bar_h), CONTENT_W, bar_h, stroke=1, fill=0)

    # Recently closed line (cumulative windows)
    _section(c, 452, "Recently closed items")
    line_top = 478
    c.setFillColor(TT_GRAY)
    c.setFont(FONT, 10.5)
    c.drawString(MARGIN, _y(line_top), "Closed in the last")
    x = MARGIN + 104
    for w in CLOSED_WINDOWS:
        c.setFont(FONT_BOLD, 16)
        c.setFillColor(TT_GRAY)
        num = str(s["closed_recent"][w])
        c.drawString(x, _y(line_top + 1), num)
        nx = x + c.stringWidth(num, FONT_BOLD, 16) + 6
        c.setFont(FONT, 10.5)
        c.setFillColor(MUTED)
        c.drawString(nx, _y(line_top), f"{w} days")
        x = nx + c.stringWidth(f"{w} days", FONT, 10.5) + 26
        if w != CLOSED_WINDOWS[-1]:
            c.setStrokeColor(LIGHT_RULE)
            c.setLineWidth(0.75)
            c.line(x - 13, _y(line_top + 4), x - 13, _y(line_top - 12))
    c.setFont(FONT, 8)
    c.setFillColor(MUTED)
    sub = "Counts are cumulative (the 60-day figure includes the last 30 days), based on Date Closed."
    if s["closed_undated"]:
        sub += f" {s['closed_undated']} closed item(s) have no Date Closed and are not counted."
    for i, line in enumerate(_wrap(c, sub, FONT, 8, CONTENT_W)):
        c.drawString(MARGIN, _y(line_top + 18 + i * 11), line)

    # Notes
    y = 536
    _section(c, y, "Notes")
    notes = [
        f"Ages are counted from each item's Observation Date to the date of distribution "
        f"({_fmt_date(s['ref_date'])}), the same rule used to highlight the Time Sensitive {log_name}.",
        "Open items include anything not marked Closed (Open and In Progress).",
        f"Closed items are excluded from the age counts: {s['closed']} of {s['total']} items in this export.",
    ]
    if s["oldest"]:
        notes.append(f"Oldest open item: {s['oldest']['days']} days, observed {_fmt_date(s['oldest']['date'])}.")
    if s["undated"]:
        notes.append(f"{s['undated']} open item(s) have no valid Observation Date and cannot be aged.")
    c.setFont(FONT, 9.5)
    y += 22
    for note in notes:
        c.setFillColor(TT_OLIVE)
        c.circle(MARGIN + 3, _y(y - 3.2), 1.6, stroke=0, fill=1)
        c.setFillColor(TT_GRAY)
        for line in _wrap(c, note, FONT, 9.5, CONTENT_W - 14):
            c.drawString(MARGIN + 12, _y(y), line)
            y += 13
        y += 4

    # TT logo, bottom-right (template: x 432–558, top 745.5)
    from PIL import Image as PILImage
    lw = 126
    pw, ph = PILImage.open(engine.LOGO_PATH).size
    lh = lw * ph / pw
    c.drawImage(str(engine.LOGO_PATH), MARGIN + CONTENT_W - lw, _y(745.5 + lh), width=lw, height=lh, mask="auto")

    c.showPage()
    c.save()
    return buf.getvalue()


def _wrap(c, text, font, size, width):
    words, lines, cur = text.split(), [], ""
    for w in words:
        trial = f"{cur} {w}".strip()
        if c.stringWidth(trial, font, size) <= width:
            cur = trial
        else:
            lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines
