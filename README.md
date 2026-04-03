# PCR Patient Care Report App

Flask + SQLite web app for creating, editing, viewing, exporting, and backing up Patient Care Report (PCR) records.

## What This App Can Do

- Create full PCR records from the form.
- Edit saved records using the same full form.
- Draw on the body diagram image (`static/images/PCR.jpg`) and save the drawing per record.
- Manage dynamic Next of Kin entries.
- View and print records.
- Delete records.
- Export all data to CSV (flattened nested fields).
- Export full database backup (`.db`) for transfer/recovery.

## Tech Stack

- Python + Flask
- SQLite (`pcr.db` by default)
- HTML/CSS/Vanilla JavaScript

## Project Structure

- [app.py](app.py): Flask routes, data collection, DB logic, export/backup endpoints.
- [templates/base.html](templates/base.html): App shell and top navigation.
- [templates/new.html](templates/new.html): Full PCR create/edit form.
- [templates/records.html](templates/records.html): Dashboard list + filters.
- [templates/view.html](templates/view.html): Printable record details.
- [static/style.css](static/style.css): Styling.
- [static/images/PCR.jpg](static/images/PCR.jpg): Body diagram image.
- `pcr.db`: SQLite DB file (auto-created).
- `backups/`: Generated DB backup files.

## Setup

1. Create and activate a virtual environment.
2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Run the app:

```bash
flask run
```

Or:

```bash
python app.py
```

4. Open in browser:

- `http://127.0.0.1:5000/records` (dashboard)
- `http://127.0.0.1:5000/new` (new PCR)

## How To Use

1. Go to Dashboard.
2. Click `New PCR`.
3. Fill out the form sections.
4. In the body diagram area:
   - `Draw`: annotate findings.
   - `Erase`: remove parts of drawing.
   - `Clear`: remove all drawing.
5. Click `Save PCR`.
6. Use Dashboard actions:
   - `View`
   - `Edit`
   - `Delete`

## Export and Backup

- Export flattened CSV:
  - Route: `/records/export.csv`
  - Includes top-level metadata and flattened nested PCR fields.
- Export full DB backup:
  - Route: `/records/export.db`
  - Also available via top nav `Backup DB`.
  - Creates a timestamped backup in `backups/` and downloads it.

## Data Persistence and Reliability

- SQLite durability settings are enabled (`WAL`, `synchronous=FULL`, `foreign_keys=ON`).
- Create/edit/delete operations include DB error handling with rollback on failure.
- If a field is empty, CSV output shows `N/A`.

## Multi-Computer Use

By default, each computer has its own local DB.

- Local DB path defaults to project `pcr.db`.
- You can set environment variable `PCR_DB_PATH` to point to a specific database file:

```powershell
$env:PCR_DB_PATH="C:\path\to\shared_or_custom\pcr.db"
flask run
```

Use DB backup export/import if you need to transfer records between computers.

## Notes

- The patient name is required.
- Next of Kin dynamic entries start at index 1.
- Injury detail rows (deformity, bleeding, contusion, tenderness, abrasion, laceration, punctured, swelling) are writable and saved.
