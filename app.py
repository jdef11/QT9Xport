"""
QT9 Document Downloader — Web Interface
Run: python app.py
Visit: http://localhost:5000
"""

import io
import re
import tempfile
import threading
import uuid
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from flask import Flask, abort, jsonify, redirect, render_template, request, send_file
from playwright.sync_api import sync_playwright

from qt9_downloader import (
    BASE_URL,
    DOCS_PAGE,
    apply_name_filter,
    apply_status_filter,
    build_session,
    download_via_http,
    login,
    sanitize_filename,
    scan_all_documents,
    screenshot,
    set_max_page_size,
    spot_check_downloads,
)

app = Flask(__name__)

DOWNLOADS_DIR = Path("./qt9_downloads/web")
TIMEOUT_MS = 30_000
DOWNLOAD_WORKERS = 5

# In-memory job store — keyed by job_id (UUID string)
# {job_id: {status, doc_refs, results, messages, error}}
JOBS: dict = {}


# ---------------------------------------------------------------------------
# Background job runner (two-phase: Playwright scan → parallel HTTP downloads)
# ---------------------------------------------------------------------------

def run_download_job(job_id: str, username: str, password: str, doc_refs: list[str]):
    job = JOBS[job_id]
    output_dir = DOWNLOADS_DIR / job_id
    shots_dir = output_dir / "screenshots"
    output_dir.mkdir(parents=True, exist_ok=True)
    shots_dir.mkdir(parents=True, exist_ok=True)

    _lock = threading.Lock()

    def push(msg: str):
        with _lock:
            job["messages"].append(msg)

    try:
        # ------------------------------------------------------------------
        # Phase 1 — Playwright: login, scan, probe URL template, grab cookies
        # ------------------------------------------------------------------
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

            push("Scanning document list…")
            docs, url_template = scan_all_documents(
                page, doc_refs, output_dir, TIMEOUT_MS, shots_dir
            )
            auth_session = build_session(context)
            browser.close()

        push(f"Scan complete: {len(docs)} document(s) found")

        # ------------------------------------------------------------------
        # Phase 2 — HTTP: parallel downloads, updating job results
        # ------------------------------------------------------------------
        if not url_template:
            job["status"] = "error"
            job["error"] = "Could not determine download URL — all matched docs may be electronic-only."
            push("ERROR: Could not determine download URL template.")
            return

        # Match scanned docs back to the requested refs
        to_download = []
        for doc in docs:
            doc_name = doc["doc_name"]
            matched_ref = None
            for ref in doc_refs:
                if doc_name.lower().startswith(ref.lower()):
                    matched_ref = ref
                    break
            if matched_ref is None:
                continue

            if doc["exists"]:
                with _lock:
                    job["results"][matched_ref]["status"] = "found"
                    job["results"][matched_ref]["files"].append(doc["existing_path"].name)
                push(f"Already on disk: {doc['existing_path'].name}")
            elif doc["doc_id"]:
                to_download.append((doc, matched_ref))
            else:
                push(f"No doc ID for: {doc_name} — skipping")

        push(f"Downloading {len(to_download)} file(s) in parallel…")

        def _http_task(doc, matched_ref):
            url = url_template.format(doc_id=doc["doc_id"])
            path, reason = download_via_http(auth_session, url, doc["doc_name"], output_dir)
            return doc, matched_ref, path, reason

        with ThreadPoolExecutor(max_workers=DOWNLOAD_WORKERS) as pool:
            futures = {
                pool.submit(_http_task, doc, ref): (doc, ref)
                for doc, ref in to_download
            }
            for future in as_completed(futures):
                try:
                    doc, matched_ref, path, reason = future.result()
                except Exception as exc:
                    doc, matched_ref = futures[future]
                    push(f"Unexpected error for '{doc['doc_name']}': {exc}")
                    with _lock:
                        job["results"][matched_ref]["status"] = "not_found"
                    continue

                with _lock:
                    if reason in ("ok", "exists"):
                        job["results"][matched_ref]["status"] = "found"
                        job["results"][matched_ref]["files"].append(path.name)
                    elif reason == "no_file":
                        job["results"][matched_ref]["status"] = "no_file"
                    else:
                        job["results"][matched_ref]["status"] = "not_found"

                if reason in ("ok", "exists"):
                    push(f"Saved: {path.name}")
                elif reason == "no_file":
                    push(f"No file stored in QT9 for: {doc['doc_name']}")
                else:
                    push(f"Download failed for: {doc['doc_name']}")

        # Any ref still pending after the scan was never found in the grid
        with _lock:
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
