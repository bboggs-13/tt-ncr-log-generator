import io
import unittest
from pathlib import Path

import openpyxl
import pandas as pd

import engine
import log_pdf

NEW_CSV = Path("/home/user/134_jane_street_export_nyc_special_inspections_-_non-conformance_log_17_11_2026 (3).csv")

ROW = {
    "Report #": "64", "Non-Conformance #": "1", "Observation Date": "2026-10-09", "Type": "TR-8",
    "Priority": "High", "Floor": "Bulkhead Roof", "Responsible Subcontractor": "TCS / Airflex",
    "Title": "Unadhered Insulation Board", "Description": "Insulation discontinuity",
    "Recommended Action": "Re-install insulation", "Required Documentation": "Sketch, Field Fix Package",
    "Status": "Open", "Updates": "Awaiting submittal", "Date Closed": "",
    "Link to ticket": "https://www.planradar.com/dr/1/2/defects/abc/details",
}
STANDARD = ["Report", "NCR #", "Observation Date", "Status", "Floor", "Responsible Subcontractor",
            "Title", "Description", "Recommended Action", "Documentation Required", "Remarks",
            "Date Closed", "Link to ticket"]
CLOSEOUT = STANDARD[:4] + ["Type"] + STANDARD[4:]


def _csv(rows):
    return pd.DataFrame(rows).to_csv(index=False).encode()


def _log(key, csv_bytes, link=True):
    xlsx, _df, msgs = engine.process_csv_bytes(csv_bytes, key, planradar_access=link)
    assert xlsx is not None, msgs
    wb = openpyxl.load_workbook(io.BytesIO(xlsx))
    ws = wb[engine.SHEET_TITLES[key]]
    return xlsx, ws, msgs


def _headers(ws):
    return [ws.cell(engine.HEADER_ROW, c).value for c in range(1, ws.max_column + 1)
            if ws.cell(engine.HEADER_ROW, c).value]


class SpecialNcrColumnTests(unittest.TestCase):
    def test_column_order_and_titles(self):
        for key, expected in (("special", STANDARD), ("special_time", STANDARD),
                              ("special_closeout", CLOSEOUT)):
            with self.subTest(key=key):
                _x, ws, _m = _log(key, _csv([ROW]))
                self.assertEqual(_headers(ws), expected)
                _x, ws, _m = _log(key, _csv([ROW]), link=False)
                self.assertEqual(_headers(ws), expected[:-1])

    def test_values_land_under_the_right_headers(self):
        _x, ws, msgs = _log("special_closeout", _csv([ROW]))
        values = dict(zip(_headers(ws), (ws.cell(8, c).value for c in range(1, 15))))
        self.assertEqual(values["NCR #"], 1)
        self.assertEqual(values["Status"], "Open")
        self.assertEqual(values["Type"], "TR-8")
        self.assertEqual(values["Floor"], "Bulkhead Roof")
        self.assertEqual(values["Responsible Subcontractor"], "TCS / Airflex")
        self.assertEqual(values["Title"], "Unadhered Insulation Board")
        self.assertEqual(values["Documentation Required"], "Sketch, Field Fix Package")
        self.assertEqual(values["Remarks"], "Awaiting submittal")      # from the CSV's Updates column
        self.assertFalse(any("left blank" in m for m in msgs))

    def test_status_rules_follow_the_moved_status_column(self):
        rows = [dict(ROW, Status=s, **{"Non-Conformance #": str(i)})
                for i, s in enumerate(["Open", "Closed", "In Progress"], start=1)]
        for key in ("special", "special_time", "special_closeout"):
            with self.subTest(key=key):
                profile = engine.get_profile(key)
                self.assertEqual(profile.column_map["Status"], "D")
                _x, ws, _m = _log(key, _csv(rows))
                self.assertEqual(ws.auto_filter.filterColumn[0].colId, 3)          # column D
                self.assertEqual(str(ws.data_validations.dataValidation[0].sqref), "D8:D10")
                self.assertTrue(ws.row_dimensions[9].hidden)                       # Closed row
                self.assertFalse(ws.row_dimensions[8].hidden)

    def test_alignment(self):
        _x, ws, _m = _log("special", _csv([ROW]))
        for col in "ABCDEFLM":
            self.assertEqual(ws[f"{col}8"].alignment.horizontal, "center", col)
        for col in "GHIJK":
            self.assertEqual(ws[f"{col}8"].alignment.horizontal, "left", col)
            self.assertTrue(ws[f"{col}8"].alignment.wrap_text, col)

    def test_legends_sit_clear_of_the_title_block(self):
        for key in ("special_time", "special_closeout"):
            with self.subTest(key=key):
                _x, ws, _m = _log(key, _csv([ROW]))
                merged = {str(m) for m in ws.merged_cells.ranges}
                self.assertIn("D1:G1", merged)
                self.assertIn("E2:G2", merged)
                self.assertTrue(str(ws["D1"].value).startswith("Legend"))

    def test_older_exports_still_work(self):
        old = dict(ROW)
        for k in ("Floor", "Responsible Subcontractor", "Title", "Required Documentation", "Updates"):
            old.pop(k)
        old["Remarks"] = "Older remark"
        _x, ws, msgs = _log("special", _csv([old]))
        values = dict(zip(_headers(ws), (ws.cell(8, c).value for c in range(1, 14))))
        self.assertEqual(values["Remarks"], "Older remark")
        self.assertIsNone(values["Floor"])
        self.assertTrue(any("left blank" in m and "'Floor'" in m for m in msgs))

    @unittest.skipUnless(NEW_CSV.exists(), "reference CSV not available")
    def test_reference_csv_and_pdf(self):
        for key in ("special", "special_time", "special_closeout"):
            with self.subTest(key=key):
                xlsx, ws, msgs = _log(key, NEW_CSV.read_bytes())
                self.assertEqual(ws["F8"].value if key != "special_closeout" else ws["G8"].value,
                                 "TCS / Airflex")
                profile = engine.with_planradar_link(engine.get_profile(key), True)
                pdf = log_pdf.build_log_pdf(xlsx, profile)
                self.assertTrue(pdf.startswith(b"%PDF"))


if __name__ == "__main__":
    unittest.main()
