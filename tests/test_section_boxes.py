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


if __name__ == "__main__":
    unittest.main()
