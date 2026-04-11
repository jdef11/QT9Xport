"""
QT9 Document Downloader — Web Interface
Run: python app.py
Visit: http://localhost:5000
"""

import io
import re
import tempfile
import threading
import time
import uuid
import zipfile
from pathlib import Path

from flask import Flask, abort, jsonify, redirect, render_template, request, send_file
from playwright.sync_api import TimeoutError as PlaywrightTimeout
from playwright.sync_api import sync_playwright

from qt9_downloader import (
    BASE_URL,
    DOCS_PAGE,
    apply_name_filter,
    apply_status_filter,
    get_grid_rows,
    get_row_doc_name,
    login,
    next_page,
    sanitize_filename,
    screenshot,
    set_max_page_size,
    spot_check_downloads,
)

app = Flask(__name__)

DOWNLOADS_DIR = Path("./qt9_downloads/web")
TIMEOUT_MS = 30_000

# In-memory job store — keyed by job_id (UUID string)
# {job_id: {status, doc_refs, results, messages, error}}
JOBS: dict = {}


# ---------------------------------------------------------------------------
# Core download helper — like qt9_downloader.download_row but returns the
# saved filepath (or None) so the web layer can track what was downloaded.
# ---------------------------------------------------------------------------

def _download_file(page, row, doc_name: str, output_dir: Path, shots_dir: Path):
    """
    Right-click a grid row and save the file.
    Returns:
        (Path, "ok")       — file saved successfully
        (Path, "exists")   — file already present in output_dir
        (None, "no_file")  — document has no file stored in QT9
        (None, "error")    — timeout or other failure
    """
    safe_name = sanitize_filename(doc_name)

    existing = list(output_dir.glob(f"{safe_name}.*"))
    if existing:
        return (existing[0], "exists")

    try:
        row.click(button="right", timeout=5000)

        try:
            page.wait_for_selector(
                "#ctl00_cphCenter_rcmCurrentDocsGridRow_detached",
                state="visible",
                timeout=5000,
            )
        except PlaywrightTimeout:
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
            return (None, "error")

        dl_link = page.query_selector(
            '#ctl00_cphCenter_rcmCurrentDocsGridRow_detached '
            'a.rmLink:has-text("Download File")'
        )
        if not dl_link or not dl_link.is_visible():
            page.keyboard.press("Escape")
            return (None, "no_file")

        with page.expect_download(timeout=TIMEOUT_MS) as dl_info:
            dl_link.click(timeout=5000)

        download = dl_info.value
        ext = Path(download.suggested_filename or f"{safe_name}.bin").suffix

        filepath = output_dir / f"{safe_name}{ext}"
        counter = 1
        while filepath.exists():
            filepath = output_dir / f"{safe_name}_{counter}{ext}"
            counter += 1

        download.save_as(str(filepath))
        return (filepath, "ok")

    except PlaywrightTimeout:
        try:
            page.keyboard.press("Escape")
        except Exception:
            pass
        return (None, "error")
    except Exception:
        try:
            page.keyboard.press("Escape")
        except Exception:
            pass
        return (None, "error")


# ---------------------------------------------------------------------------
# Background job runner
# ---------------------------------------------------------------------------

def run_download_job(job_id: str, username: str, password: str, doc_refs: list[str]):
    job = JOBS[job_id]
    output_dir = DOWNLOADS_DIR / job_id
    shots_dir = output_dir / "screenshots"
    output_dir.mkdir(parents=True, exist_ok=True)
    shots_dir.mkdir(parents=True, exist_ok=True)

    def push(msg: str):
        job["messages"].append(msg)

    try:
        with tempfile.TemporaryDirectory(prefix="qt9_tmp_") as tmp_dir, \
                sync_playwright() as p:
            browser = p.chromium.launch(headless=True, downloads_path=tmp_dir)
            context = browser.new_context(
                accept_downloads=True,
                viewport={"width": 1400, "height": 900},
            )
            page = context.new_page()

            push("Logging in…")
            if not login(page, BASE_URL, username, password, TIMEOUT_MS, shots_dir):
                job["status"] = "error"
                job["error"] = "Login failed — check your credentials."
                return

            push("Navigating to document list…")
            page.goto(BASE_URL + DOCS_PAGE, timeout=TIMEOUT_MS)
            page.wait_for_load_state("networkidle", timeout=TIMEOUT_MS)

            apply_status_filter(page, "All/Any", TIMEOUT_MS, shots_dir)
            set_max_page_size(page, TIMEOUT_MS, shots_dir)
            apply_name_filter(page, doc_refs, TIMEOUT_MS, shots_dir)

            # Track which refs still need to be found (case-insensitive)
            remaining = set(r.lower() for r in doc_refs)
            page_num = 1

            while True:
                push(f"Scanning page {page_num}…")
                rows = get_grid_rows(page)
                if not rows:
                    break

                i = 0
                while i < len(rows):
                    doc_name = get_row_doc_name(rows[i])

                    matched_ref = None
                    for ref in doc_refs:
                        if doc_name.lower().startswith(ref.lower()):
                            matched_ref = ref
                            break

                    if not matched_ref:
                        i += 1
                        continue

                    push(f"Found: {doc_name} — downloading…")
                    job["results"][matched_ref]["status"] = "downloading"

                    filepath, reason = _download_file(
                        page, rows[i], doc_name, output_dir, shots_dir
                    )

                    if reason in ("ok", "exists"):
                        job["results"][matched_ref]["status"] = "found"
                        job["results"][matched_ref]["files"].append(filepath.name)
                        remaining.discard(matched_ref.lower())
                        push(f"Saved: {filepath.name}")
                    elif reason == "no_file":
                        job["results"][matched_ref]["status"] = "no_file"
                        remaining.discard(matched_ref.lower())
                        push(f"No file stored in QT9 for: {doc_name}")
                    else:
                        push(f"Download failed for: {doc_name}")
                        # Re-fetch only after a failure in case the DOM shifted
                        rows = get_grid_rows(page)
                        if i >= len(rows):
                            push(f"Row {i} disappeared after failed download — stopping page scan")
                            break
                        continue  # retry same index with fresh row reference

                    i += 1

                if not remaining:
                    push("All documents processed.")
                    break

                if not next_page(page, TIMEOUT_MS):
                    break
                page_num += 1

            browser.close()

        # Any ref still pending/downloading after full scan was not found
        for ref in doc_refs:
            if job["results"][ref]["status"] in ("pending", "downloading"):
                job["results"][ref]["status"] = "not_found"

        passed = spot_check_downloads(output_dir)
        if not passed:
            push("WARNING: file integrity check found suspicious files — see server logs")

        job["status"] = "done"
        push("Complete.")

    except Exception as exc:
        job["status"] = "error"
        job["error"] = str(exc)
        push(f"Unexpected error: {exc}")


# ---------------------------------------------------------------------------
# Flask routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/start", methods=["POST"])
def start():
    raw = request.form.get("doc_refs", "")
    username = request.form.get("username", "").strip()
    password = request.form.get("password", "")

    doc_refs = [s.strip() for s in re.split(r"[,\n\r]+", raw) if s.strip()]

    if not doc_refs or not username or not password:
        return redirect("/")

    job_id = str(uuid.uuid4())
    JOBS[job_id] = {
        "status": "running",
        "doc_refs": doc_refs,
        "results": {ref: {"status": "pending", "files": []} for ref in doc_refs},
        "messages": [],
        "error": None,
    }

    t = threading.Thread(
        target=run_download_job,
        args=(job_id, username, password, doc_refs),
        daemon=True,
    )
    t.start()

    return redirect(f"/job/{job_id}")


@app.route("/job/<job_id>")
def job_page(job_id):
    if job_id not in JOBS:
        abort(404)
    return render_template("job.html", job_id=job_id)


@app.route("/api/job/<job_id>")
def job_api(job_id):
    if job_id not in JOBS:
        abort(404)
    return jsonify(JOBS[job_id])


@app.route("/files/<job_id>/<filename>")
def serve_file(job_id, filename):
    # Prevent path traversal
    if ".." in filename or "/" in filename or "\\" in filename:
        abort(400)
    filepath = DOWNLOADS_DIR / job_id / filename
    if not filepath.exists():
        abort(404)
    return send_file(str(filepath.resolve()), as_attachment=True)


@app.route("/zip/<job_id>")
def serve_zip(job_id):
    job = JOBS.get(job_id)
    if not job:
        abort(404)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for ref_data in job["results"].values():
            for fname in ref_data.get("files", []):
                fpath = DOWNLOADS_DIR / job_id / fname
                if fpath.exists():
                    zf.write(str(fpath.resolve()), fname)
    buf.seek(0)
    return send_file(
        buf,
        mimetype="application/zip",
        as_attachment=True,
        download_name="qt9_documents.zip",
    )


if __name__ == "__main__":
    DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
    app.run(host="0.0.0.0", port=5000, debug=False)
