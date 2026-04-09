# QT9Xport

Automates document downloads from a QT9 QMS instance using Playwright. Two entry points:

- **`qt9_downloader.py`** — CLI script; bulk-downloads all documents matching name prefixes
- **`app.py`** — Flask web app; lets users download specific documents by reference number via a browser UI

## Setup

```bash
pip install -r requirements.txt
python -m playwright install chromium
```

## Usage

### Web app

```bash
python app.py
```

Open `http://localhost:5000`, enter document reference numbers and your QT9 credentials, and download files individually or as a ZIP.

### CLI

```bash
# Download all QMS/SDS docs (prompts for credentials)
python qt9_downloader.py

# Show browser window (useful for debugging)
python qt9_downloader.py --headed

# Filter by name prefix (comma-separated)
python qt9_downloader.py --name-prefix "QMSD-0411,SDS-032"

# Download everything
python qt9_downloader.py --name-prefix ""

# Custom output directory
python qt9_downloader.py --output ./my_docs
```

## Output

```
qt9_downloads/
  web/<job-uuid>/      # web app downloads, one folder per job
  *.pdf / *.docx ...   # CLI downloads land here
  logs/run_YYYYMMDD_HHMMSS.log
  screenshots/HHMMSS_label.png
```

Screenshots and logs are written at every key step and on every error — check these first when debugging.
