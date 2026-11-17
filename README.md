[README.md](https://github.com/user-attachments/files/34992776/README.md)
# Site Visit Post-Processing Tools

A Thornton Tomasetti web app for finishing PlanRadar site visit outputs. After the title, the app asks **What would you like to do?** and runs one of two tools:


| Tool                    | Upload                               | Download                                                                        |
| ----------------------- | ------------------------------------ | ------------------------------------------------------------------------------- |
| **Post-process report** | Site visit report (PDF)              | Post-processed copy of the PDF (captions, section borders, spacing, signatures) |
| **Generate an NCR log** | PlanRadar non-conformance CSV export | Formatted TT Excel log (+ summary PDF for Time Sensitive logs)                  |


## Post-process report

* Edits only pages headed **Photographs:**; every other page is left as is
* Links each photo to the observation on the preceding detail page (NCR or GO)
* Photos are reviewed one observation at a time: each observation ticket (e.g. `67.NCR.1`, `67.GO.1`) gets one card with small previews of its photos, laid out in the same grid as the PDF
* Photos can be reordered within their observation: pick a new position for a photo and the others shift along. Positions run from the top left across each row (photo 1 top left, photo 2 top right, photo 3 left side of row 2, and so on), including onto the next page when an observation's photos run over. **Reset photo order** puts them back
* Each photo keeps its own description when it moves; its photo number becomes its new position
* Caption format: `REPORT.NCR.OBS-n: detail description` (or `GO` for general observations), where `n` is the photo number
* Captions are left-aligned under each photo; if any caption needs a smaller font to clear the next photo row, every caption uses that same size
* Observation type and number are set once per observation; each photo's description can be edited
* Optional: section borders, spacing clean-up and signatures
* The original PDF is never overwritten
* **Non-conformance summary email** (after the captioned PDF is ready): a ready-to-send email built from the report — subject line plus, for each non-conformance, its subject, status, responsible contractor, location, description, recommended action and required documentation, and the corrected items from the Executive Summary
  * **Copy email** copies it with its formatting for pasting into Outlook; **Copy subject** copies the subject line; **Download as Outlook draft (.eml)** opens as a new unsent message in Outlook for Windows/Mac
  * Details come from each NCR's "Non-Conformance Observation" page first, then the Summary of Observations / General Observations list, then Action Items
  * **Check before sending** lists report issues it spotted (e.g. an Executive Summary NCN # that doesn't match the report, identical recommended actions, NCRs without a detail page). These are never put in the email
  * Greeting, opening sentence, closing and sign-off can be edited under **Edit wording**

Engines: `caption_engine.py` (captions, photo order, signatures) and `section_boxes.py` (borders, spacing), both PyMuPDF. Tests: `tests/`.

## Generate an NCR log

## What it does automatically

* Copies the master template (`assets/TTFE_Template.xlsx`) — the master is never modified
* Special Inspections NCR logs (Standard, Time Sensitive, Project Closeout) use these columns, starting at row 8:

  | Column | CSV column |
  |---|---|
  | Report | Report # |
  | NCR # | Non-Conformance # |
  | Observation Date | Observation Date |
  | Status | Status |
  | Type (Project Closeout log only) | Type |
  | Floor | Floor |
  | Responsible Subcontractor | Responsible Subcontractor |
  | Title | Title |
  | Description | Description |
  | Recommended Action | Recommended Action |
  | Documentation Required | Required Documentation |
  | Remarks | Updates (or Remarks in older exports) |
  | Date Closed | Date Closed |
  | Link to ticket | Link to ticket (only with PlanRadar access) |

  Priority is not shown. Older exports without Floor, Responsible Subcontractor, Title or
  Required Documentation still work; those columns are left blank and the app says so.
* Regular Site Inspection logs keep their own columns (Report, Item, Observation Date, Responsible Party, Description, Recommended Action, Status, Remarks, Date Closed)
* The engine writes the column titles (row 7) itself from the same mapping it uses for the data, and checks them before saving, so a title can never sit over another column's data. `app.py` and `engine.py` carry matching version numbers (shown under the downloads as "NCR log engine …"); if they don't match, the app stops and asks you to upload both and reboot
* Sets the project name (A4) and today's date (C5)
* Highlights the Status cell: Open = `#BDD7EE`, In Progress = `#FFE699`, Closed = `#C6E0B4`
* Centres the short columns (Report, NCR #, Observation Date, Status, Type, Floor, Responsible Subcontractor, Date Closed, Link to ticket); long-text columns are left aligned and wrapped
* Log PDF: landscape with 0.5 in margins — 11x17 for the Special Inspections NCR logs, US Letter for the Site Inspection logs
* Places one 3-inch TT logo in the top corner
* Adds an Import Audit sheet
* Time Sensitive logs: **Generate summary PDF** button at the bottom of the app — a one-page TT-style summary (project, date of distribution, open items in each age category: &lt;30, 30–59, 60–89, 90+ days) using the same rules as the log highlighting

## Project layout

```
ttfe_ncr_webapp/
├── app.py            # Streamlit UI: tool choice, then post-process report or NCR log
├── caption_engine.py # photo detection, grouping, reordering, captions, signatures (no UI)
├── section_boxes.py  # section borders and spacing clean-up (no UI)
├── ncr_email.py      # non-conformance summary email from the report (no UI)
├── engine.py         # NCR log import logic + fixed mapping (no UI)
├── summary_pdf.py    # Time Sensitive one-page summary PDF (reportlab)
├── requirements.txt
├── assets/
│   ├── TTFE_Template.xlsx
│   └── tt_logo.png
├── tests/
│   ├── test_caption_engine.py
│   ├── test_photo_order.py
│   └── test_section_boxes.py
└── README.md
```

## Run locally

```bash
cd ttfe_ncr_webapp
python -m pip install -r requirements.txt
streamlit run app.py
```

Then open the URL Streamlit prints (default http://localhost:8501).

## Deploy

Any host that runs a Python web process works. Options:

* **Streamlit Community Cloud** — push this folder to a Git repo, point the app at `app.py`.
* **Internal server / VM** — `pip install -r requirements.txt` then `streamlit run app.py --server.port 8501 --server.address 0.0.0.0`.
* **Docker** — base image `python:3.12-slim`, copy the folder, install requirements, `CMD ["streamlit","run","app.py","--server.address=0.0.0.0"]`.

## Updating the template or mapping

* Replace `assets/TTFE_Template.xlsx` to change the master layout.
* Edit the constants at the top of `engine.py` (`COLUMN_MAP`, `HEADER_ROW`, `DATA_START_ROW`, fill colors, `CENTER_COLUMNS`) if the format changes. No UI code changes are needed for a template/mapping update.

