# PCR BATO

Flask + SQLite Patient Care Report (PCR) system with web and desktop modes.

## Features

- Create, edit, view, print, and delete PCR records
- Full multi-section PCR form (assessment, vitals, team, consent/refusal, narrative)
- Body diagram drawing (draw/erase/clear) saved per record
- Dynamic Next of Kin and Crew member entries
- CSV export of records
- XLSX logsheet export using the project template
- Database backup export (`.db`)
- Dashboard filtering and search
- Optional desktop app wrapper (`dist/PCR_BATO_Desktop.exe`)

## Tech Stack

- Python 3
- Flask
- SQLite
- openpyxl
- pywebview + waitress (desktop mode)

## Current Folder Layout

- `app.py` - main Flask app, DB logic, exports, routes
- `desktop_app.py` - desktop launcher wrapper
- `templates/` - UI templates
- `static/` - CSS, service worker, images
- `instance/logsheet_template.xlsx` - XLSX template used for export
- `instance/README.md` - instance notes
- `dist/PCR_BATO_Desktop.exe` - built desktop executable
- `requirements.txt` - Python dependencies

## How to Run (App Version)

1. Create and activate a virtual environment

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Start the app:

```bash
python app.py
```

Or run via Flask:

```bash
flask --app app run
```

4. Open:

- `http://127.0.0.1:5000/records`

## Run (Desktop)

- Launch `dist/PCR_BATO_Desktop.exe`

The desktop wrapper starts a local server and opens the app in a desktop window.

## Main Routes

- `/records` - dashboard
- `/new` - create record
- `/records/<id>` - view record
- `/records/<id>/edit` - edit record
- `/records/<id>/delete` - delete record
- `/records/delete-all` - delete all records
- `/records/export.csv` - CSV export
- `/records/export.xlsx` - XLSX export
- `/records/export.db` - DB backup export

## XLSX Export Columns (Current Order)

1. Type of Emergency
2. Date of Incident
3. Time of Incident
4. Location of Incident
5. Chief Complaint
6. Patient Name
7. Age
8. Gender
9. Address
10. Contact Number
11. Driver
12. Responders
13. Communicator
14. Remarks

## Data Storage

- **Source mode**: data directory defaults to `instance/`
- **Desktop EXE mode**: data directory defaults to `%LOCALAPPDATA%/PCR_BATO`

Configurable environment variables:

- `PCR_BATO_DATA_DIR` - override app data directory
- `PCR_DB_PATH` - override database file path
- `PCR_SEED_SAMPLE_DATA` - `1`/`0` to enable/disable demo seed data

Example:

```powershell
$env:PCR_BATO_DATA_DIR = "C:\PCR_BATO_Data"
$env:PCR_DB_PATH = "C:\PCR_BATO_Data\pcr.db"
python app.py
```

## Notes

- DB uses durability settings: `WAL`, `synchronous=FULL`, `foreign_keys=ON`
- `patient_name` is required
- Export and backup are accessible from the top navigation
