import sys
import unittest
from pathlib import Path

import fitz
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import engine  # noqa: E402
import log_pdf  # noqa: E402


def _csv(rows):
    return pd.DataFrame(rows).to_csv(index=False).encode()


SPECIAL = [
    {"Report #": "1", "Non-Conformance #": "1", "Observation Date": "2020-01-01", "Type": "TR-1",
     "Description": "Old open item " * 20, "Recommended Action": "Fix it", "Status": "Open",
     "Remarks": "", "Date Closed": "", "Link to ticket": "https://www.planradar.com/dr/1/2/defects/abc/details"},
    {"Report #": "2", "Non-Conformance #": "1", "Observation Date": "2020-02-01", "Type": "QA/QC",
     "Description": "Closed item", "Recommended Action": "Done", "Status": "Closed",
     "Remarks": "", "Date Closed": "2020-03-01", "Link to ticket": ""},
    {"Report #": "3", "Non-Conformance #": "2", "Observation Date": "2020-02-03", "Type": "TR-8",
     "Description": "In progress item", "Recommended Action": "Working", "Status": "In Progress",
     "Remarks": "Note", "Date Closed": "", "Link to ticket": ""},
]


def _pdf(key, rows=SPECIAL, link=True):
    xlsx, _df, msgs = engine.process_csv_bytes(_csv(rows), inspection_type=key, planradar_access=link)
    assert xlsx is not None, msgs
    profile = engine.with_planradar_link(engine.get_profile(key), link)
    return xlsx, profile, fitz.open(stream=log_pdf.build_log_pdf(xlsx, profile), filetype="pdf")


class LogPdfTests(unittest.TestCase):
    def test_every_log_type_is_landscape_letter_with_half_inch_margins(self):
        for key in engine.PROFILES:
            with self.subTest(key=key):
                _x, _p, doc = _pdf(key)
                for page in doc:
                    self.assertEqual((page.rect.width, page.rect.height), (792, 612))
                    drawings = [d["rect"] for d in page.get_drawings()]
                    self.assertAlmostEqual(min(r.x0 for r in drawings), 36, places=1)
                    self.assertAlmostEqual(max(r.x1 for r in drawings), 756, places=1)
                    blocks = page.get_text("blocks")
                    self.assertGreaterEqual(min(b[1] for b in blocks), 35.5)
                    self.assertLessEqual(max(b[3] for b in blocks), 612 - 35.5)

    def test_shows_only_the_rows_the_workbook_shows(self):
        _x, _p, doc = _pdf("special")
        text = "".join(page.get_text() for page in doc)
        self.assertIn("Old open item", text)
        self.assertIn("In progress item", text)
        self.assertNotIn("Closed item", text)
        self.assertIn("2 open items shown", text)

    def test_full_link_text_and_column_dropped_without_planradar_access(self):
        _x, _p, doc = _pdf("special")
        flat = "".join(doc[0].get_text().split())
        self.assertIn("https://www.planradar.com/dr/1/2/defects/abc/details", flat)
        _x, _p, doc = _pdf("special", link=False)
        self.assertNotIn("Link to ticket", doc[0].get_text())

    def test_colours_match_the_log(self):
        x, p, _doc = _pdf("special_time")
        log = log_pdf.read_log(x, p)
        red = log_pdf._rgb(engine.AGE_BANDS[-1][3])
        self.assertTrue(all(f is not None and f.hexval() == red.hexval() for f in log["row_fills"]))
        x, p, _doc = _pdf("special_closeout")
        log = log_pdf.read_log(x, p)
        tr1 = log_pdf._rgb(engine.TYPE_FILLS[0][2])
        self.assertEqual(log["row_fills"][0].hexval(), tr1.hexval())
        x, p, _doc = _pdf("special")
        log = log_pdf.read_log(x, p)
        self.assertEqual(log["cell_fills"][0].hexval(), log_pdf._rgb(engine.OPEN_FILL).hexval())
        self.assertEqual(log["cell_fills"][1].hexval(), log_pdf._rgb(engine.INPROGRESS_FILL).hexval())

    def test_long_logs_continue_with_headers_repeated(self):
        rows = [dict(SPECIAL[0], **{"Non-Conformance #": str(i)}) for i in range(60)]
        _x, _p, doc = _pdf("special", rows=rows)
        self.assertGreater(len(doc), 1)
        for page in doc:
            self.assertIn("Recommended Action", page.get_text())
            self.assertIn(f"of {len(doc)}", page.get_text())

    def test_no_open_items_still_makes_a_pdf(self):
        rows = [dict(SPECIAL[1])]
        _x, _p, doc = _pdf("regular", rows=rows)
        self.assertIn("No open items.", doc[0].get_text())


if __name__ == "__main__":
    unittest.main()
