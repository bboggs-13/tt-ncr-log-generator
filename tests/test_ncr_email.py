import email
import unittest
from pathlib import Path

import fitz

import ncr_email as ne

HOME = Path("/home/user")
R67 = HOME / "Report_134_Jane_Street_3_0_NYC___Special_Inspections_Report_Template_03_11_2026_16_16_42.pdf"
RAW65 = HOME / "261029_PlanRadar_SI Report_RAW.pdf"


def _report(pages):
    """Small synthetic report: each page is a list of lines."""
    doc = fitz.open()
    for lines in pages:
        page = doc.new_page(width=612, height=792)
        y = 40
        for line in ["TT Project No:N1.01", "Inspection Date: 01/02/2026", "Report Date: 01/05/2026",
                     "Special Inspection Report #12"] + lines:
            page.insert_text((36, y), line, fontsize=8)
            y += 11
    return doc.tobytes()


SYNTH = _report([
    ["A. Executive Summary", "Project: Sample Tower", "GC/CM: Builder Co", "TT Inspector: Jane Doe",
     "Any Non-Conformances Observed: Yes", "NCN #: 12.NC.1, 12.NC.2",
     "Any Non-Conformances Corrected: No", "NCN#:"],
    ["C. Non-Conformance Observation 12.NC.1", "TT Inspector: Jane Doe", "Contractor: Glazing Sub <A&B>",
     "Status: Open", "Subject: Missing end dams", "Description: End dams were not installed",
     "at the sill flashing.", "Recommended Action: Install end dams.",
     "Required Documentation: Photos", "Location:", "Floor: 3", "Elevation: North", "Grid Lines: N/A"],
    ["D. Summary of Observations", "Non-conformance Observations:",
     "12.NC.1 End dams were not installed at the sill flashing.", "Recommended Action(s):",
     "Install end dams.", "12.NC.2 Sealant joint width below minimum.", "Recommended Action(s):",
     "Remove and replace sealant.", "General Observations:", "1. Work in progress."],
])


class ExtractTests(unittest.TestCase):
    def test_synthetic_report(self):
        s = ne.extract_summary(SYNTH)
        self.assertEqual((s.project, s.report_no, s.inspection_date), ("Sample Tower", "12", "01/02/2026"))
        self.assertEqual([it.id for it in s.items], ["12.NC.1"])     # 12.NC.2 is only in the summary
        one = s.items[0]
        self.assertEqual(one.subject, "Missing end dams")
        self.assertEqual(one.contractor, "Glazing Sub <A&B>")
        self.assertEqual(one.description, "End dams were not installed at the sill flashing.")
        self.assertEqual(one.location, "Floor 3, North Elevation")
        self.assertEqual(one.recommended_action, "Install end dams.")
        self.assertEqual(s.corrected_ids, [])
        notes = ne.check_notes(s)
        self.assertTrue(any("lists 12.NC.1, 12.NC.2" in n for n in notes))   # exec summary vs ticket pages

    @unittest.skipUnless(R67.exists(), "report #67 not available")
    def test_report_67(self):
        s = ne.extract_summary(R67.read_bytes())
        self.assertEqual((s.project, s.report_no), ("134 Jane Street", "67"))
        self.assertEqual([it.id for it in s.items], ["67.NC.1"])     # the only ticket page
        first = s.items[0]
        self.assertEqual(first.subject, "Unadhered Insulation Board Behind Aluminum Rainscreen")
        self.assertEqual((first.status, first.contractor), ("Open", "TCS / Airflex"))
        self.assertEqual(first.location, "Bulkhead Roof, South Elevation")
        self.assertTrue(first.required_documentation.startswith("Sketch, Field Fix Package"))
        self.assertEqual(s.corrected_ids, ["22.NC.3", "53.NC.1", "55.NC.1", "58.NC.4"])
        notes = ne.check_notes(s)
        self.assertTrue(any("lists 66.NC.1" in n for n in notes))

    @unittest.skipUnless(R67.exists(), "report #67 not available")
    def test_report_without_summary_section(self):
        """Reports no longer have a summary at the end: the email must not change."""
        with_summary = ne.extract_summary(R67.read_bytes())
        doc = fitz.open(stream=R67.read_bytes(), filetype="pdf")
        doc.delete_pages(6, 7)                        # E. Summary of Observations
        without = ne.extract_summary(doc.tobytes())
        self.assertEqual(ne.build_email(without).text, ne.build_email(with_summary).text)
        self.assertEqual(without.items[0].description, with_summary.items[0].description)

    @unittest.skipUnless(RAW65.exists(), "raw report not available")
    def test_action_items_layout(self):
        s = ne.extract_summary(RAW65.read_bytes())
        self.assertEqual([it.id for it in s.items], ["65.NC.1"])     # Action Items list is not read
        self.assertEqual(s.items[0].sources, ["detail page"])


class EmailTests(unittest.TestCase):
    def test_email_content_and_escaping(self):
        em = ne.build_email(ne.extract_summary(SYNTH))
        self.assertEqual(em.subject, "Sample Tower \u2013 Special Inspection Report #12 \u2013 Non-Conformance Summary")
        self.assertIn("<b>12.NC.1 \u2013 Missing end dams</b>", em.html)
        self.assertIn("Glazing Sub &lt;A&amp;B&gt;", em.html)          # HTML-escaped
        self.assertIn("Responsible: Glazing Sub <A&B>", em.text)
        self.assertNotIn("verified as corrected", em.text)            # none corrected
        self.assertIn("Hello all,", em.text)

    def test_wording_options(self):
        s = ne.extract_summary(SYNTH)
        s.corrected_ids = ["3.NC.1"]
        self.assertEqual(s.inspection_type, "Special Inspection")
        self.assertEqual(ne.default_intro(s),
                         "Please find the attached Special Inspection Report #12 and associated NCR log for "
                         "TT's visit on Friday, 01/02/2026. Documented below is a summary of the "
                         "non-conformances observed during TT's special inspection.")
        self.assertEqual(ne._with_weekday("11/10/2026"), "Tuesday, 11/10/2026")
        self.assertEqual(ne._with_weekday("10/30/26"), "Friday, 10/30/26")
        self.assertEqual(ne._with_weekday("TBD"), "TBD")
        em = ne.build_email(s, ne.EmailWording(greeting="Hi Sal,", intro="Custom intro.", closing="",
                                               sign_off="Best,", include_corrected=False))
        self.assertTrue(em.text.startswith("Hi Sal,\n\nCustom intro."))
        self.assertNotIn("3.NC.1", em.text)
        self.assertTrue(em.text.rstrip().endswith("Best,"))

    def test_eml_is_an_unsent_draft(self):
        em = ne.build_email(ne.extract_summary(SYNTH))
        msg = email.message_from_bytes(ne.build_eml(em))
        self.assertEqual(msg["X-Unsent"], "1")
        self.assertEqual([p.get_content_type() for p in msg.walk()],
                         ["multipart/alternative", "text/plain", "text/html"])

    def test_report_without_ncrs(self):
        s = ne.extract_summary(_report([["A. Executive Summary", "Project: Sample Tower",
                                         "Any Non-Conformances Observed: No", "NCN #: N/A"]]))
        self.assertEqual(s.items, [])
        self.assertIn("No non-conformances were observed", ne.build_email(s).text)


if __name__ == "__main__":
    unittest.main()
