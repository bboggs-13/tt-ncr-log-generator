[README.md](https://github.com/user-attachments/files/32915291/README.md)
# TT NCR Log Generator

A self-contained web app that converts a PlanRadar Non-Conformance /
Special-Inspections CSV export into the standard Thornton Tomasetti Excel log.
The template, TT logo, field mapping, Status highlighting, and column
centering are all built in — the user only uploads a CSV and downloads the
finished workbook.

## What it does automatically

- Copies the master template (`assets/TTFE_Template.xlsx`) — the master is never modified
- Writes 11 CSV columns into template columns A–K starting at row 8
- Sets the project name (A4) and today's date (C5)
- Highlights the Status cell: Open = `#BDD7EE`, Closed = `#C6E0B4`
- Centers Report #, Item, Observation Date, Type, Priority, Status, Date Closed, Link to ticket
- Places one 3-inch TT logo in the top corner
- Adds an Import Audit sheet

## Project layout

```
ttfe_ncr_webapp/
├── app.py            # Streamlit UI (upload → validate → preview → download)
├── engine.py         # reusable import logic + fixed mapping (no UI)
├── requirements.txt
├── assets/
│   ├── TTFE_Template.xlsx
│   └── tt_logo.png
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
