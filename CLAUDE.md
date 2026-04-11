# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project does

Two entry points that both automate document downloads from a QT9 QMS instance (`https://oxfordpm.qt9qms.app`) using Playwright:

- **`qt9_downloader.py`** — CLI script; bulk-downloads all documents matching name prefixes
- **`app.py`** — Flask web app; lets users download specific documents by reference number via a browser UI

## Running

```bash
# One-time setup
pip install -r requirements.txt
python -m playwright install chromium

# Web app (share http://localhost:5000 with coworkers)
python app.py

# CLI — prompts for credentials, downloads all QMS/SDS docs
python qt9_downloader.py

# CLI options
python qt9_downloader.py --headed          # show browser (debugging)
python qt9_downloader.py --name-prefix "QMSD-0411,SDS-032"
python qt9_downloader.py --name-prefix ""  # download everything
python qt9_downloader.py --output ./my_docs
```

## Output structure

```
qt9_downloads/
  web/<job-uuid>/          # web app downloads, one folder per job
  *.pdf / *.docx ...       # CLI downloads land directly here
  logs/run_YYYYMMDD_HHMMSS.log
  screenshots/HHMMSS_label.png
```

Screenshots and logs are written at every key step and on every error — check these first when debugging.

## Architecture

### Shared automation layer — `qt9_downloader.py`

All Playwright logic lives here. Both the CLI and web app use these functions:

```
login()
  ├─ fill_login_form()   # ctl00_cphCenter_txtUserName / txtPassword / btnSubmit
  └─ is_logged_in()      # polls up to 20s for logout button or nav links
apply_status_filter()    # select#ctl00_cphCenter_ddlStatus
set_max_page_size()      # <select> page-size dropdown (RadComboBox not supported)
apply_name_filter()      # input[id*="FilterTextBox_DocumentName"]; skipped for multiple prefixes
get_grid_rows()          # tr.rgRow / tr.rgAltRow
get_row_doc_name()       # skips display:none TDs; first visible TD is DocumentName
download_row()           # right-click → #rcmCurrentDocsGridRow_detached → a.rmLink
next_page()              # .rgPageNext if not disabled
```

`main()` in `qt9_downloader.py` is only called by the CLI. `app.py` imports the functions above directly.

### Web app — `app.py`

Flask server with in-memory job store (`JOBS` dict). Flow:

```
POST /start
  └─ creates JOBS[job_id], starts background thread → run_download_job()
       ├─ _download_file()   # local variant of download_row(); returns (Path, reason)
       │                       reason: "ok" | "exists" | "no_file" | "error"
       └─ updates JOBS[job_id]["results"] per ref as it processes rows

GET /job/<id>             → job.html (polls /api/job/<id> every 1.5s via JS)
GET /api/job/<id>         → JSON snapshot of JOBS[job_id]
GET /files/<id>/<fname>   → serves downloaded file
GET /zip/<id>             → builds and streams ZIP of all job files
```

**Matching logic in `run_download_job`**: doc refs from the form are matched case-insensitively with `doc_name.lower().startswith(ref.lower())`. No server-side grid filter is applied — all pages are iterated. Stops early once all refs are resolved.

**Result statuses**: `pending` → `downloading` → `found` | `no_file` | `not_found`

### Templates — `templates/`

- `index.html` — form: textarea for ref numbers, username, password
- `job.html` — progress page; pure JS polling, no framework

## Key QT9 quirks

- **Login page**: `/Default.aspx` — takes ~5s for fields to become enabled after load
- **Post-login redirect**: `networkidle` fires mid-redirect through `/Login.aspx`; success is detected by polling for DOM indicators (`input[id*="LogOut"]`, nav links), not by URL
- **Hidden form fields**: Registration fields (`firstName`, `lastName`, etc.) persist in the DOM after login — don't use `input[type="password"]` presence as a failure signal
- **Grid row layout**: TD[0] and TD[1] are `display:none` (status flags and numeric doc ID). TD[2] (first visible) is `DocumentName`
- **Context menu**: Telerik `RadContextMenu` ID `ctl00_cphCenter_rcmCurrentDocsGridRow`. After right-click, the `_detached` container becomes visible. Menu items use `a.rmLink`
- **"Download File" visibility**: QT9's `SetMenuItems()` hides the item when `getDataKeyValue("Electronic") == 'True'` — but JS type coercion means this never fires; hidden = no file stored
- **Playwright downloads**: `downloads_path` on `browser.launch()` is set to a `tempfile.TemporaryDirectory` — keeps Playwright staging files out of the output folder. The temp dir is auto-cleaned when the `with` block exits. `download.save_as()` is still used to move each file to its final location.
- **Page size dropdown**: Telerik `RadComboBox`, not a standard `<select>` — `set_max_page_size()` won't find it; pages default to 50 rows
