[README.md](https://github.com/user-attachments/files/34092250/README.md)
# Site Visit Post-Processing Tools

A Thornton Tomasetti web app for finishing PlanRadar site visit outputs. After the
title, the app asks **What would you like to do?** and runs one of two tools:

| Tool | Upload | Download |
|---|---|---|
| **Generate report captions** | Post-processed site visit report (PDF) | Captioned copy of the PDF |
| **Generate an NCR log** | PlanRadar non-conformance CSV export | Formatted TT Excel log (+ summary PDF for Time Sensitive logs) |

## Generate report captions

- Edits only pages headed **Photographs:**; every other page is left as is
- Links each photo to the observation on the preceding detail page (NCR or GO)
- Caption format: `REPORT.NCR.OBS-n: existing caption` (or `GO` for general observations), where `n` is the photo number
- Captions are left-aligned under each photo; if any caption needs a smaller font to clear the next photo row, every caption uses that same size
- Each photo can be reviewed (type, observation number, caption text) before export
- The original PDF is never overwritten

Engine: `caption_engine.py` (PyMuPDF). Tests: `tests/test_caption_engine.py`.

## Generate an NCR log

## What it does automatically

- Copies the master template (`assets/TTFE_Template.xlsx`) — the master is never modified
- Writes 11 CSV columns into template columns A–K starting at row 8
- Sets the project name (A4) and today's date (C5)
- Highlights the Status cell: Open = `#BDD7EE`, Closed = `#C6E0B4`
- Centers Report #, Item, Observation Date, Status, Date Closed, Link to ticket (plus Type on the Project Closeout log)
- Places one 3-inch TT logo in the top corner
- Adds an Import Audit sheet
- Time Sensitive logs: **Generate summary PDF** button at the bottom of the app — a one-page TT-style summary (project, date of distribution, open items in each age category: <30, 30–59, 60–89, 90+ days) using the same rules as the log highlighting

## Project layout

```
ttfe_ncr_webapp/
├── app.py            # Streamlit UI: tool choice, then captions or NCR log
├── caption_engine.py # photo-caption detection and PDF editing (no UI)
├── engine.py         # NCR log import logic + fixed mapping (no UI)
├── summary_pdf.py    # Time Sensitive one-page summary PDF (reportlab)
├── requirements.txt
├── assets/
│   ├── TTFE_Template.xlsx
│   └── tt_logo.png
├── tests/
│   └── test_caption_engine.py
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

- **Streamlit Community Cloud** — push this folder to a Git repo, point the app at `app.py`.
- **Internal server / VM** — `pip install -r requirements.txt` then
  `streamlit run app.py --server.port 8501 --server.address 0.0.0.0`.
- **Docker** — base image `python:3.12-slim`, copy the folder, install
  requirements, `CMD ["streamlit","run","app.py","--server.address=0.0.0.0"]`.

## Updating the template or mapping

- Replace `assets/TTFE_Template.xlsx` to change the master layout.
- Edit the constants at the top of `engine.py` (`COLUMN_MAP`, `HEADER_ROW`,
  `DATA_START_ROW`, fill colors, `CENTER_COLUMNS`) if the format changes.
No UI code changes are needed for a template/mapping update.
