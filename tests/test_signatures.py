import io
import unittest
from pathlib import Path

import fitz
from PIL import Image, ImageDraw

import caption_engine as ce

REPORT = Path("/home/user/Report_134_Jane_Street_3_0_NYC___Special_Inspections_Report_Template_29_10_2026_13_21_26.pdf")


def signature_png(width=900, height=300, paper=(244, 241, 232), transparent=False):
    image = Image.new("RGBA", (width, height), (255, 255, 255, 0) if transparent else paper + (255,))
    draw = ImageDraw.Draw(image)
    draw.line([(width * 0.1, height * 0.6), (width * 0.3, height * 0.3), (width * 0.5, height * 0.7),
               (width * 0.8, height * 0.4)], fill=(20, 40, 110, 255), width=8)
    out = io.BytesIO()
    image.save(out, format="PNG")
    return out.getvalue()


def sign_off_page_pdf():
    """A page with the report's sign-off row: [Signed: ........ ][Date: 10/29/2026]."""
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    grey = (0.87, 0.87, 0.87)
    for rect in (fitz.Rect(28.5, 294.75, 228.75, 312.75), fitz.Rect(228.75, 294.75, 347.25, 312.75)):
        page.draw_rect(rect, color=None, fill=(1, 1, 1))
        page.draw_rect(rect, color=grey, width=0.7)
    page.insert_text((30, 307.6), "Signed:", fontname="hebo", fontsize=9)
    page.insert_text((230.25, 307.6), "Date: 10/29/2026", fontname="helv", fontsize=9)
    return doc.tobytes()


class SignatureTests(unittest.TestCase):
    def test_finds_area_between_label_and_date_box(self):
        box = ce.find_signature_box(sign_off_page_pdf())
        self.assertIsNotNone(box)
        area, cell, date = fitz.Rect(box.area), fitz.Rect(box.cell), fitz.Rect(box.date_cell)
        self.assertGreater(area.x0, 62)            # right of "Signed:"
        self.assertLess(area.x1, date.x0)          # stops before the Date box
        self.assertTrue(cell.contains(area))       # inside the Signed cell

    def test_one_and_two_signatures_fit_inside_and_do_not_overlap(self):
        area = fitz.Rect(69, 297.5, 226, 310.75)
        for ratios in ([3.0], [3.0, 3.0], [0.7, 6.0], [12.0, 12.0]):
            with self.subTest(ratios=ratios):
                rects = ce.layout_signatures(area, ratios)
                self.assertEqual(len(rects), len(ratios))
                for rect, ratio in zip(rects, ratios):
                    self.assertTrue(area.contains(rect))
                    self.assertAlmostEqual(rect.width / rect.height, ratio, places=3)   # not distorted
                if len(rects) == 2:
                    self.assertLess(rects[0].x1, rects[1].x0)                         # side by side

    def test_background_is_removed_and_margins_trimmed(self):
        png, ratio = ce.prepare_signature(signature_png())
        image = Image.open(io.BytesIO(png))
        self.assertEqual(image.getpixel((0, 0))[3], 0)            # paper is transparent
        self.assertLess(image.width, 900)                         # blank margins trimmed
        self.assertAlmostEqual(ratio, image.width / image.height)

    def test_blank_image_is_rejected(self):
        blank = io.BytesIO()
        Image.new("RGB", (400, 200), "white").save(blank, format="PNG")
        with self.assertRaises(ValueError):
            ce.prepare_signature(blank.getvalue())

    def test_apply_places_two_images_in_the_box(self):
        output, warnings, info = ce.apply_signatures(
            sign_off_page_pdf(), [signature_png(), signature_png(transparent=True)])
        self.assertEqual(warnings, [])
        self.assertEqual(info["placed"], 2)
        page = fitz.open(stream=output, filetype="pdf")[0]
        area, date = fitz.Rect(info["area"]), fitz.Rect(info["date_cell"])
        placed = [fitz.Rect(i["bbox"]) for i in page.get_image_info()]
        self.assertEqual(len(placed), 2)
        for rect in placed:
            self.assertTrue(area.contains(rect))
            self.assertLess(rect.x1, date.x0)

    def test_more_than_two_is_refused(self):
        output, warnings, info = ce.apply_signatures(sign_off_page_pdf(), [signature_png()] * 3)
        self.assertEqual(info["placed"], 0)
        self.assertTrue(warnings)

    def test_report_without_signed_box_is_left_unchanged(self):
        doc = fitz.open(); doc.new_page()
        source = doc.tobytes()
        output, warnings, info = ce.apply_signatures(source, [signature_png()])
        self.assertEqual(output, source)
        self.assertEqual(info["placed"], 0)

    @unittest.skipUnless(REPORT.exists(), "134 Jane Street sample report not available")
    def test_real_report_signature_box(self):
        box = ce.find_signature_box(REPORT.read_bytes())
        self.assertEqual(box.page_number, 4)
        self.assertLess(fitz.Rect(box.area).x1, fitz.Rect(box.date_cell).x0)


if __name__ == "__main__":
    unittest.main()
