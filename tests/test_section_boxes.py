import sys
import unittest
from pathlib import Path

import fitz

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import section_boxes as sb  # noqa: E402

GREY = (0xDD / 255,) * 3


def _page(doc, top=84.8, bottom=771.8):
    page = doc.new_page(width=595, height=842)
    page.draw_rect(fitz.Rect(28.5, top, 566.2, bottom), color=None, fill=(1, 1, 1))   # body area
    return page


def make_raw_report():
    """B/C/D sections laid out like PlanRadar's raw export (no boxes, text at 28.5 pt)."""
    doc = fitz.open()
    p1 = _page(doc)
    p1.insert_text((28.5, 120), "B. General Observations", fontsize=10.5, fontname="hebo")
    p1.insert_text((28.5, 150), "Non-conformance Observations:", fontsize=9, fontname="hebo")
    p1.insert_text((43.5, 168), "65.NC.1 A long observation line that runs right up to the margin", fontsize=9)
    p1.insert_text((83.2, 750), "a wrapped line near the page bottom", fontsize=9)
    p2 = _page(doc)
    p2.insert_text((83.2, 96), "continued on the next page", fontsize=9)
    p2.insert_text((28.5, 140), "Work-in-Progress Observations:", fontsize=9, fontname="hebo")
    p2.insert_text((69.0, 158), "1. Installation in progress", fontsize=9)
    p2.insert_text((28.5, 200), "C. Follow-Up Items", fontsize=10.5, fontname="hebo")
    p2.insert_text((28.5, 230), "The following items are to be verified:", fontsize=9, fontname="hebo")
    p2.insert_text((69.0, 248), "1. Verify flashing", fontsize=9)
    p2.insert_text((28.5, 300), "E. Non-Conformance Observation 65.NC.1", fontsize=10.5, fontname="hebo")
    p2.insert_text((30.0, 330), "Location:", fontsize=9, fontname="hebo")   # not in B/C/D
    return doc.tobytes()


def make_new_format_report():
    """Newer format: B. Follow-Up Items, C/D ticket pages, then E. Summary of Observations."""
    doc = fitz.open()
    p1 = _page(doc)
    p1.insert_text((28.5, 120), "B. Follow-Up Items", fontsize=10.5, fontname="hebo")
    p1.insert_text((28.5, 150), "The following items are to be verified:", fontsize=9, fontname="hebo")
    p1.insert_text((69.0, 168), "1. Verify flashing", fontsize=9)
    p2 = _page(doc)
    p2.insert_text((28.5, 120), "C. Non-Conformance Observation 67.NC.1", fontsize=10.5, fontname="hebo")
    p2.insert_text((30.0, 150), "Location:", fontsize=9, fontname="hebo")
    p2.insert_text((28.5, 300), "D. General Observation #1", fontsize=10.5, fontname="hebo")
    p2.insert_text((30.0, 330), "Location:", fontsize=9, fontname="hebo")
    p3 = _page(doc)
    p3.insert_text((28.5, 120), "E. Summary of Observations", fontsize=10.5, fontname="hebo")
    p3.insert_text((28.5, 150), "Non-conformance Observations:", fontsize=9, fontname="hebo")
    p3.insert_text((43.5, 168), "67.NC.1 A discontinuity in the thermal insulation", fontsize=9)
    p3.insert_text((82.5, 184), "Recommended Action(s):", fontsize=9)
    p3.insert_text((31.0, 700), "   General Observations:", fontsize=9, fontname="hebo")
    p3.insert_text((69.0, 750), "1. Installation of roofing substrates", fontsize=9)
    p4 = _page(doc)
    p4.insert_text((69.0, 110), "2. Installation of aluminum cladding", fontsize=9)
    p4.insert_text((28.5, 140), "Note: All conditions were reviewed with the GC.", fontsize=9)
    p4.insert_text((28.5, 200), "F. Appendix", fontsize=10.5, fontname="hebo")
    p4.insert_text((30.0, 230), "Attendees:", fontsize=9, fontname="hebo")
    return doc.tobytes()


def _boxes(pdf, page_index):
    page = fitz.open(stream=pdf, filetype="pdf")[page_index]
    return sb._grey_rules(page)


class SectionBoxTests(unittest.TestCase):
    def setUp(self):
        self.src = make_raw_report()
        self.result = sb.apply_section_boxes(self.src)

    def test_boxes_only_subsections_of_b_c_d(self):
        self.assertEqual(self.result.sections, ["Non-conformance Observations:",
                                                "Work-in-Progress Observations:",
                                                "The following items are to be verified:"])
        kinds = [(b.page_index, b.kind) for b in self.result.boxes]
        self.assertEqual(kinds, [(0, "heading"), (0, "content"), (1, "content"),
                                 (1, "heading"), (1, "content"), (1, "heading"), (1, "content")])

    def test_heading_cell_is_18pt_and_content_starts_under_it(self):
        head, content = self.result.boxes[0], self.result.boxes[1]
        self.assertAlmostEqual(head.rect[3] - head.rect[1], 18.0, places=2)
        self.assertEqual(content.rect[1], head.rect[3])
        self.assertEqual((head.rect[0], head.rect[2]), (28.5, 565.5))

    def test_box_reopens_at_top_of_next_page_and_meets_next_subheading(self):
        cont = self.result.boxes[2]
        self.assertTrue(cont.open_top)
        self.assertAlmostEqual(cont.rect[1], 84.8, places=1)
        self.assertAlmostEqual(cont.rect[3], self.result.boxes[3].rect[1], places=2)

    def test_main_headings_are_not_boxed(self):
        for pdf_page, heading_y in ((0, 120), (1, 200)):
            for rule in _boxes(self.result.pdf, pdf_page):
                self.assertFalse(rule.y0 - 1 < heading_y - 4 < rule.y1 + 1 and rule.width > 100,
                                 f"a border crosses the main heading on page {pdf_page + 1}")

    def test_borders_use_report_grey_and_width(self):
        rules = _boxes(self.result.pdf, 0)
        self.assertTrue(rules)
        self.assertTrue(all(min(r.width, r.height) <= sb.LINE + 0.01 for r in rules))

    def test_text_is_inset_unchanged_and_not_duplicated(self):
        before = fitz.open(stream=self.src, filetype="pdf")
        after = fitz.open(stream=self.result.pdf, filetype="pdf")
        for a, b in zip(before, after):
            self.assertEqual(a.get_text(), b.get_text())
        page = after[0]
        head = page.search_for("Non-conformance Observations:")[0]
        self.assertAlmostEqual(head.x0, 30.0, places=1)
        right = max(c["bbox"][2] for b in page.get_text("rawdict")["blocks"] for l in b.get("lines", [])
                    for s in l["spans"] for c in s["chars"] if c["c"].strip())
        self.assertLess(right, sb.RIGHT - sb.LINE)
        main = page.search_for("B. General Observations")[0]
        self.assertAlmostEqual(main.x0, 28.5, places=1)               # outside the boxes: unmoved
        loc = after[1].search_for("Location:")[0]
        self.assertAlmostEqual(loc.x0, 30.0, places=1)

    def test_running_twice_adds_nothing(self):
        again = sb.apply_section_boxes(self.result.pdf)
        self.assertEqual(again.boxes, [])
        self.assertEqual(again.pdf, self.result.pdf)


class NewFormatTests(unittest.TestCase):
    """Sections are found by name, so the letters can move between report formats."""

    def setUp(self):
        self.result = sb.apply_section_boxes(make_new_format_report())

    def test_boxes_follow_up_and_summary_of_observations_only(self):
        self.assertEqual(self.result.sections, ["The following items are to be verified:",
                                                "Non-conformance Observations:",
                                                "General Observations:"])
        self.assertEqual({b.page_index for b in self.result.boxes}, {0, 2, 3})   # ticket pages untouched

    def test_summary_box_continues_onto_next_page_and_stops_before_appendix(self):
        cont = [b for b in self.result.boxes if b.page_index == 3]
        self.assertEqual(len(cont), 1)
        self.assertTrue(cont[0].open_top)
        self.assertGreater(cont[0].rect[3], 140)            # the Note stays inside the box
        self.assertLess(cont[0].rect[3], 200 - 10)          # F. Appendix stays outside

    def test_running_twice_adds_nothing(self):
        self.assertEqual(sb.apply_section_boxes(self.result.pdf).boxes, [])


def make_gappy_report(photo=False):
    """Summary section with one oversized gap and one list item with no gap."""
    doc = fitz.open()
    p = _page(doc)
    p.insert_text((28.5, 120), "E. Summary of Observations", fontsize=10.5, fontname="hebo")
    p.draw_rect(fitz.Rect(28.5, 122, 54.8, 125), color=None, fill=(0.83, 0.27, 0.11))   # orange rule
    p.insert_text((28.5, 150), "Non-conformance Observations:", fontsize=9, fontname="hebo")
    y = 168
    for n in (1, 2, 3):                       # standard spacing: 22.2 pt between items
        p.insert_text((43.5, y), f"67.NC.{n} Observation text", fontsize=9)
        y += 32.2
    p.insert_text((28.5, y + 60), "General Observations:", fontsize=9, fontname="hebo")   # oversized gap
    g = y + 78
    p.insert_text((69.0, g), "1. Installation of roofing", fontsize=9)
    p.draw_line(fitz.Point(69, g + 2), fitz.Point(150, g + 2), color=(0.2, 0.2, 0.2), width=0.5)
    p.insert_text((69.0, g + 32.2), "2. Installation of cladding", fontsize=9)
    p.insert_text((69.0, g + 48.0), "3. Installation of brick", fontsize=9)     # no gap above
    p.insert_text((69.0, g + 80.2), "4. Installation of windows", fontsize=9)
    if photo:
        pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 40, 30), 0)
        pix.clear_with(0x999999)
        p.insert_image(fitz.Rect(300, 600, 400, 680), pixmap=pix)
    return doc.tobytes()


def _tops(pdf, *needles):
    page = fitz.open(stream=pdf, filetype="pdf")[0]
    return [page.search_for(n)[0].y0 for n in needles]


class TightenSpacingTests(unittest.TestCase):
    def test_off_by_default(self):
        src = make_gappy_report()
        self.assertEqual(sb.apply_section_boxes(src).moves, [])

    def test_large_gap_closes_box_like_the_others_and_items_are_even(self):
        src = make_gappy_report()
        r = sb.apply_section_boxes(src, tighten=True)
        self.assertEqual(len(r.moves), 2)
        nc3, head, i1, i2, i3, i4 = _tops(r.pdf, "67.NC.3", "General Observations:", "1. Installation",
                                          "2. Installation", "3. Installation", "4. Installation")
        content = [b for b in r.boxes if b.kind == "content"][0]
        heading2 = [b for b in r.boxes if b.kind == "heading"][1]
        self.assertAlmostEqual(content.rect[3], heading2.rect[1], places=1)
        self.assertLess(head - nc3, 30)                                  # was ~100 pt
        self.assertAlmostEqual(i2 - i1, i3 - i2, delta=0.3)              # evened out
        self.assertAlmostEqual(i3 - i2, i4 - i3, delta=0.3)

    def test_graphics_move_with_their_text_and_words_are_unchanged(self):
        src = make_gappy_report()
        r = sb.apply_section_boxes(src, tighten=True, boxes=False)

        def offset(pdf):
            page = fitz.open(stream=pdf, filetype="pdf")[0]
            item1 = page.search_for("1. Installation")[0]
            rule = [d["rect"] for d in page.get_drawings() if d.get("color") and d["rect"].width > 50][0]
            return item1.y1, rule.y0 - item1.y1

        (y_before, off_before), (y_after, off_after) = offset(src), offset(r.pdf)
        self.assertLess(y_after, y_before - 30)                          # the item moved up
        self.assertAlmostEqual(off_after, off_before, places=2)          # its underline came with it
        page = fitz.open(stream=r.pdf, filetype="pdf")[0]
        before = sorted(w[4] for w in fitz.open(stream=src, filetype="pdf")[0].get_text("words"))
        self.assertEqual(before, sorted(w[4] for w in page.get_text("words")))
        orange = [d["rect"] for d in page.get_drawings()
                  if d.get("fill") and d["fill"][0] > 0.8 and d["fill"][1] < 0.5][0]
        self.assertAlmostEqual(orange.y0, 122, places=1)                 # above the gaps: unmoved

    def test_pages_with_photos_are_left_alone(self):
        r = sb.apply_section_boxes(make_gappy_report(photo=True), tighten=True)
        self.assertEqual(r.moves, [])
        self.assertTrue(any("photo" in w for w in r.warnings))

    def test_finished_report_spacing_is_left_alone(self):
        r = sb.apply_section_boxes(make_raw_report(), tighten=True)
        again = sb.apply_section_boxes(r.pdf, tighten=True)
        self.assertEqual(again.moves, [])


if __name__ == "__main__":
    unittest.main()
