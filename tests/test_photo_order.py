"""Photo grouping, previews and reordering within an observation's photo grid."""
import unittest

import fitz

import caption_engine as ce

COLS = (33.75, 300.75)
ROWS = (102.0, 318.75, 535.5)
W, H = 259.5, 194.25
SHADES = [0x10, 0x28, 0x40, 0x58, 0x70, 0x88, 0xA0, 0xB8, 0xD0]     # one grey level per photo


def _photo_page(doc, photos):
    page = doc.new_page(width=595, height=842)
    page.insert_text((30, 62), "Special Inspection Report #70", fontsize=13.5)
    page.insert_text((30, 98), "Photographs:", fontsize=9)
    for slot, (shade, caption) in enumerate(photos):
        x, y = COLS[slot % 2], ROWS[slot // 2]
        pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 80, 60), 0)
        pix.clear_with(shade)
        page.insert_image(fitz.Rect(x, y, x + W, y + H), pixmap=pix)
        page.insert_text((x, y + H + 10.5), caption, fontsize=10.5)


def make_report():
    """NCR 1 with 2 photos; GO 1 with 8 photos: 6 on one page and 2 on the next."""
    doc = fitz.open()
    p = doc.new_page(width=595, height=842)
    p.insert_text((30, 62), "Special Inspection Report #70", fontsize=13.5)
    p.insert_text((30, 110), "E. Non-Conformance Observation 70.NC.1", fontsize=12)
    _photo_page(doc, [(SHADES[0], "NCR photo A"), (SHADES[1], "NCR photo B")])
    p = doc.new_page(width=595, height=842)
    p.insert_text((30, 62), "Special Inspection Report #70", fontsize=13.5)
    p.insert_text((30, 110), "F. General Observation # 1", fontsize=12)
    _photo_page(doc, [(SHADES[i % len(SHADES)], f"GO photo {i + 1}") for i in range(6)])
    _photo_page(doc, [(SHADES[6], "GO photo 7"), (SHADES[7], "GO photo 8")])
    return doc.tobytes()


def _colour_at(pdf, page_index, rect):
    page = fitz.open(stream=pdf, filetype="pdf")[page_index]
    r = fitz.Rect(rect)
    pix = page.get_pixmap(clip=fitz.Rect(r.x0 + 20, r.y0 + 20, r.x0 + 21, r.y0 + 21))
    return pix.pixel(0, 0)[0]


class PhotoGroupTests(unittest.TestCase):
    def setUp(self):
        self.src = make_report()
        self.report, self.jobs, _ = ce.analyze_pdf(self.src)
        self.groups = ce.group_photos(self.jobs)

    def test_photos_are_grouped_by_observation_in_grid_order(self):
        self.assertEqual([g.key for g in self.groups], ["NCR.1", "GO.1"])
        self.assertEqual([len(g.jobs) for g in self.groups], [2, 8])
        self.assertEqual(self.groups[1].pages, [4, 5])
        self.assertEqual(self.groups[1].columns, 2)
        self.assertEqual(self.groups[0].title(self.report), "70.NCR.1 \u00b7 Non-Conformance Observation 1")
        numbers = [self.jobs[i].photo_number for i in self.groups[1].jobs]
        self.assertEqual(numbers, list(range(1, 9)))

    def test_previews_are_small_images_of_each_photo(self):
        previews = ce.photo_previews(self.src, self.jobs, width=120)
        self.assertEqual(len(previews), len(self.jobs))
        first = fitz.Pixmap(previews[0])
        self.assertAlmostEqual(first.width, 120, delta=2)

    def test_reorder_moves_photos_and_their_descriptions_to_new_slots(self):
        go = self.groups[1].jobs                       # job indexes in slot order
        order = [go[7], go[0], go[1], go[2], go[3], go[4], go[5], go[6]]   # last photo first
        out, new_jobs, notes = ce.reorder_photos(self.src, self.jobs, {"GO.1": order}, self.groups)
        self.assertIn("Photo order changed", notes[0])
        # Slot 1 (page 4, top left) now shows the photo that was on page 5, slot 2.
        self.assertEqual(_colour_at(out, 3, self.jobs[go[0]].image_rect), SHADES[7])
        self.assertEqual(_colour_at(out, 3, self.jobs[go[1]].image_rect), SHADES[0])
        self.assertEqual(_colour_at(out, 4, self.jobs[go[6]].image_rect), SHADES[5 % len(SHADES)])
        self.assertEqual(_colour_at(out, 4, self.jobs[go[7]].image_rect), SHADES[6])
        # Descriptions follow their photos; photo numbers follow the grid.
        self.assertEqual([new_jobs[i].caption for i in go[:3]], ["GO photo 8", "GO photo 1", "GO photo 2"])
        self.assertEqual([new_jobs[i].photo_number for i in go], list(range(1, 9)))
        # Same number of photos on every page; nothing else changed.
        before, after = fitz.open(stream=self.src, filetype="pdf"), fitz.open(stream=out, filetype="pdf")
        for a, b in zip(before, after):
            self.assertEqual(len(a.get_image_info()), len(b.get_image_info()))
            self.assertEqual(a.get_text(), b.get_text())

    def test_captions_after_reorder(self):
        go = self.groups[1].jobs
        order = [go[1], go[0]] + go[2:]
        out, new_jobs, _ = ce.reorder_photos(self.src, self.jobs, {"GO.1": order}, self.groups)
        pdf, warnings = ce.apply_captions(out, new_jobs, self.report)
        lines = [l for l in fitz.open(stream=pdf, filetype="pdf")[3].get_text().splitlines() if l.startswith("70.")]
        self.assertEqual(lines[:2], ["70.GO.1-1: GO photo 2", "70.GO.1-2: GO photo 1"])

    def test_unchanged_order_returns_the_original_pdf(self):
        out, new_jobs, notes = ce.reorder_photos(self.src, self.jobs, {g.key: g.jobs for g in self.groups},
                                                 self.groups)
        self.assertIs(out, self.src)
        self.assertEqual(notes, [])

    def test_invalid_order_is_ignored(self):
        out, _, notes = ce.reorder_photos(self.src, self.jobs, {"GO.1": [0, 0, 1]}, self.groups)
        self.assertIs(out, self.src)
        self.assertTrue(notes)


if __name__ == "__main__":
    unittest.main()
