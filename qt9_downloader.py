"""
QT9 QMS Bulk Document Downloader
Downloads all current documents from https://oxfordpm.qt9qms.app/CurrentDocuments.aspx
by right-clicking each grid row and selecting "Download File".

Usage:
    python qt9_downloader.py
    python qt9_downloader.py --output ./downloads
    python qt9_downloader.py --headed
    python qt9_downloader.py --filter Active

Options:
    --output    Download folder (default: ./qt9_downloads)
    --headed    Show browser window (useful for debugging)
    --timeout   Page load timeout in seconds (default: 30)
    --filter    Status filter to apply: Active, All/Any, etc. (default: All/Any)
    --url       Base URL override (default: https://oxfordpm.qt9qms.app)
"""

import argparse
import getpass
import logging
import mimetypes
import random
import re
import sys
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

import requests
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout


BASE_URL = "https://oxfordpm.qt9qms.app"
DOCS_PAGE = "/CurrentDocuments.aspx"

# Module-level logger — configured in main()
log = logging.getLogger("qt9")

# Throttle context-menu screenshots to the first N documents only
_ctx_screenshot_count = 0
_CTX_SCREENSHOT_LIMIT = 2


def setup_logging(log_dir: Path) -> Path:
    """Write logs to both console and a timestamped file."""
    log_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"run_{ts}.log"

    fmt = logging.Formatter("%(asctime)s  %(levelname)-7s  %(message)s", datefmt="%H:%M:%S")

    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(fmt)

    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)

    log.setLevel(logging.DEBUG)
    log.addHandler(fh)
    log.addHandler(ch)

    return log_file


def screenshot(page, label: str, shots_dir: Path):
    """Save a full-page screenshot with a timestamped name."""
    ts = datetime.now().strftime("%H%M%S")
    name = f"{ts}_{label}.png"
    path = shots_dir / name
    try:
        page.screenshot(path=str(path), full_page=True)
        log.debug(f"Screenshot → {path.name}")
    except Exception as e:
        log.warning(f"Screenshot failed ({label}): {e}")


def parse_args():
    parser = argparse.ArgumentParser(description="QT9 QMS Bulk Document Downloader")
    parser.add_argument("--url", default=BASE_URL, help="Base URL of the QT9 instance")
    parser.add_argument("--output", default="./qt9_downloads", help="Download folder")
    parser.add_argument("--headed", action="store_true", help="Show browser window")
    parser.add_argument("--timeout", type=int, default=30, help="Timeout in seconds")
    parser.add_argument("--filter", default="All/Any", dest="status_filter",
                        help="Status filter (default: All/Any)")
    parser.add_argument("--name-prefix", default="QMS,SDS", dest="name_prefix",
                        help="Comma-separated prefixes — only download docs whose name starts with "
                             "one of these (default: QMS,SDS). Set to empty string to download all.")
    parser.add_argument("--workers", type=int, default=5,
                        help="Parallel HTTP download workers (default: 5)")
    return parser.parse_args()


def prompt_credentials() -> tuple[str, str]:
    print("QT9 QMS Credentials")
    print("-" * 20)
    username = input("Username: ").strip()
    password = getpass.getpass("Password: ")
    print()
    return username, password


def sanitize_filename(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*\r\n\t]', "_", name)
    name = re.sub(r'\s+', " ", name).strip()
    return name[:200]


def dismiss_modal(page):
    """Close any overlay/popup that may block interaction."""
    modal_close_selectors = [
        'button:has-text("Close")',
        'button:has-text("OK")',
        'button:has-text("×")',
        '.modal-close',
        '[aria-label="Close"]',
    ]
    for sel in modal_close_selectors:
        try:
            el = page.query_selector(sel)
            if el and el.is_visible():
                el.click(timeout=2000)
                log.info(f"Dismissed modal via '{sel}'")
                time.sleep(0.5)
                return
        except Exception:
            continue

    # If a visible overlay div is blocking, try pressing Escape
    try:
        page.keyboard.press("Escape")
    except Exception:
        pass


def fill_login_form(page, username: str, password: str, shots_dir: Path) -> bool:
    """Fill and submit the QT9 login form on whatever page it appears."""
    inputs = page.evaluate("""() =>
        Array.from(document.querySelectorAll('input'))
            .filter(i => i.offsetParent !== null)
            .map(i => ({id: i.id, name: i.name, type: i.type}))
    """)
    log.debug(f"Visible inputs: {inputs}")

    # Dismiss any modal that may be covering the form
    dismiss_modal(page)

    # Username — try specific known IDs first, then fallbacks
    user_filled = False
    for sel in [
        'input[id="ctl00_cphCenter_txtUserName"]',
        'input[name="ctl00$cphCenter$txtUserName"]',
        'input[name*="UserName"]',
        'input[id*="UserName"]',
        'input[name="username"]',
    ]:
        try:
            el = page.query_selector(sel)
            if el and el.is_visible():
                el.click()
                el.fill(username)
                actual = el.input_value()
                if actual == username:
                    log.info(f"Username filled — selector: '{sel}'")
                    user_filled = True
                    break
                else:
                    log.warning(f"'{sel}' readback mismatch: got '{actual}'")
        except Exception as e:
            log.debug(f"Username selector '{sel}' failed: {e}")

    # Password
    pass_filled = False
    for sel in [
        'input[id="ctl00_cphCenter_txtPassword"]',
        'input[name="ctl00$cphCenter$txtPassword"]',
        'input[name*="Password"]',
        'input[id*="Password"]',
        'input[type="password"]',
    ]:
        try:
            el = page.query_selector(sel)
            if el and el.is_visible():
                el.click()
                el.fill(password)
                pass_filled = True
                log.info(f"Password filled — selector: '{sel}'")
                break
        except Exception as e:
            log.debug(f"Password selector '{sel}' failed: {e}")

    screenshot(page, "03_login_filled", shots_dir)

    if not user_filled or not pass_filled:
        log.error(f"Could not fill fields — user_filled={user_filled}, pass_filled={pass_filled}")
        screenshot(page, "03_login_fill_failed", shots_dir)
        return False

    # Submit — use specific known ID first
    submitted = False
    for sel in [
        'input[id="ctl00_cphCenter_btnSubmit_input"]',
        'input[name="ctl00$cphCenter$btnSubmit"]',
        'input[type="submit"]',
        'button[type="submit"]',
        'button:has-text("Log In")',
        'input[value*="Log"]',
    ]:
        try:
            el = page.query_selector(sel)
            if el and el.is_visible():
                el.click()
                submitted = True
                log.info(f"Form submitted — selector: '{sel}'")
                break
        except Exception as e:
            log.debug(f"Submit selector '{sel}' failed: {e}")

    if not submitted:
        page.keyboard.press("Enter")
        log.info("Form submitted via Enter key")

    return True


def is_logged_in(page) -> str | None:
    """Return the matching selector if authenticated indicators are found, else None."""
    for sel in [
        'input[id*="LogOut"]',
        'img[id*="LogOut"]',
        'a:has-text("Logout")',
        'a:has-text("Doc. Control")',
        'a:has-text("ISO Functions")',
    ]:
        try:
            if page.query_selector(sel):
                return sel
        except Exception:
            continue
    return None


def login(page, base_url: str, username: str, password: str,
          timeout_ms: int, shots_dir: Path) -> bool:
    login_url = base_url.rstrip("/") + "/Default.aspx"
    log.info(f"Navigating to login page: {login_url}")
    page.goto(login_url, timeout=timeout_ms)
    page.wait_for_load_state("networkidle", timeout=timeout_ms)
    screenshot(page, "01_login_page", shots_dir)

    log.info("Waiting for login fields to become enabled...")
    try:
        page.wait_for_function(
            "() => { const el = document.querySelector('input[id*=\"txtUserName\"]'); "
            "return el && !el.disabled; }",
            timeout=10000,
        )
    except PlaywrightTimeout:
        log.debug("Field-enable wait timed out — falling back to 3s sleep")
        time.sleep(3)
    screenshot(page, "02_login_fields_ready", shots_dir)

    if not fill_login_form(page, username, password, shots_dir):
        return False

    # Poll for authenticated indicators for up to 20 seconds.
    # QT9 redirects through Login.aspx before landing on the home page,
    # and networkidle can fire mid-redirect before the final page renders.
    log.info("Waiting for authenticated page to load (up to 20s)...")
    deadline = time.time() + 20
    matched = None
    while time.time() < deadline:
        try:
            page.wait_for_load_state("networkidle", timeout=3000)
        except PlaywrightTimeout:
            pass
        matched = is_logged_in(page)
        if matched:
            break
        log.debug(f"Not authenticated yet — URL: {page.url} — retrying...")
        time.sleep(1)

    screenshot(page, "04_post_login", shots_dir)
    log.info(f"URL after submit: {page.url}")

    if matched:
        log.info(f"Login successful — authenticated indicator: '{matched}' — {page.url}")
        return True

    log.error("Login failed — no authenticated indicators found after 20s")
    screenshot(page, "04_login_failed", shots_dir)
    return False


def apply_status_filter(page, status_filter: str, timeout_ms: int, shots_dir: Path):
    try:
        # Try specific known ID first (ctl00_cphCenter_ddlStatus)
        dropdown = page.query_selector('select#ctl00_cphCenter_ddlStatus')
        if not dropdown:
            for d in page.query_selector_all("select"):
                opts = [o.inner_text().strip() for o in d.query_selector_all("option")]
                if any("All" in o or "Active" in o for o in opts):
                    dropdown = d
                    break

        if dropdown:
            opts = [o.inner_text().strip() for o in dropdown.query_selector_all("option")]
            log.debug(f"Select options found: {opts}")
            dropdown.select_option(label=status_filter)
            _wait_for_grid(page, timeout_ms)
            log.info(f"Status filter set to: {status_filter}")
            screenshot(page, "06_filter_applied", shots_dir)
        else:
            log.warning("Status filter dropdown not found — proceeding with default")
    except Exception as e:
        log.warning(f"Status filter skipped: {e}")


def set_max_page_size(page, timeout_ms: int, shots_dir: Path):
    try:
        # Try legacy native <select> first (older QT9 versions)
        select = page.query_selector(
            "select.rgPageSizeDD, select[id*='PageSize'], select[id*='pageSize']"
        )
        if select:
            options = select.query_selector_all("option")
            max_val, max_num = None, 0
            for opt in options:
                try:
                    n = int(opt.get_attribute("value") or opt.inner_text())
                    if n > max_num:
                        max_num, max_val = n, opt.get_attribute("value") or opt.inner_text()
                except ValueError:
                    pass
            if max_val:
                select.select_option(max_val)
                _wait_for_grid(page, timeout_ms)
                log.info(f"Grid page size set to {max_val} rows (native select)")
                screenshot(page, "07_page_size_max", shots_dir)
            return

        # Telerik RadComboBox: click the arrow, find the highest li, click it
        arrow_btn = page.query_selector(
            "a[id*='PageSizeComboBox'][id*='Arrow'], "
            ".RadComboBox[id*='PageSizeComboBox'] .rcbArrowCell"
        )
        if not arrow_btn:
            log.debug("Page size control not found (neither native <select> nor RadComboBox)")
            return

        arrow_btn.click(timeout=5000)
        try:
            page.wait_for_selector(
                ".RadComboBox[id*='PageSizeComboBox'] .rcbList li",
                state="visible",
                timeout=5000,
            )
        except PlaywrightTimeout:
            log.debug("RadComboBox page-size list did not open")
            return

        items = page.query_selector_all(
            ".RadComboBox[id*='PageSizeComboBox'] .rcbList li"
        )
        max_item, max_num = None, 0
        for item in items:
            try:
                n = int(item.inner_text().strip())
                if n > max_num:
                    max_num, max_item = n, item
            except ValueError:
                pass

        if max_item:
            max_item.click(timeout=5000)
            _wait_for_grid(page, timeout_ms)
            log.info(f"Grid page size set to {max_num} rows (RadComboBox)")
            screenshot(page, "07_page_size_max", shots_dir)

    except Exception as e:
        log.debug(f"set_max_page_size: {e}")


def apply_name_filter(page, prefixes: list, timeout_ms: int, shots_dir: Path):
    """
    Use the Telerik RadGrid built-in filter row to restrict rows to those
    whose document-name column starts with `prefix`.

    Telerik filter rows live in <tr class="rgFilterRow">.  We find the input
    in the column that corresponds to the document name (skip the first
    boolean/checkbox column) and type the prefix, then press Enter.
    If the filter row isn't found we log a warning and fall back to the
    per-row skip guard in the download loop.
    """
    if not prefixes:
        return

    # The grid filter only supports a single term. With multiple prefixes we skip
    # it and rely entirely on the per-row check in the download loop.
    if len(prefixes) > 1:
        log.info(f"Multiple prefixes {prefixes} — skipping grid filter, using per-row check")
        return

    prefix = prefixes[0]

    try:
        # Try specific known filter input ID first
        target = page.query_selector('input[id*="FilterTextBox_DocumentName"]')
        if not target:
            filter_row = page.query_selector("tr.rgFilterRow")
            if not filter_row:
                log.warning("Telerik filter row (tr.rgFilterRow) not found — "
                            "will skip non-matching rows in the download loop instead")
                return
            filter_inputs = filter_row.query_selector_all("td input[type='text']")
            if not filter_inputs:
                log.warning("No text inputs found in filter row — skipping grid filter")
                return
            # First column is hidden boolean/status; DocumentName filter is second input
            target = filter_inputs[1] if len(filter_inputs) > 1 else filter_inputs[0]

        target.click()
        target.fill(prefix)
        log.debug(f"Typed '{prefix}' into grid name-filter input")

        # Some Telerik grids need the filter type set to "StartsWith".
        # Try to find the nearby filter-type button/dropdown and set it.
        try:
            # The filter type menu button is usually a sibling element with
            # class rgFilterTypeButton or similar.
            filter_td = target.evaluate_handle("el => el.closest('td')")
            type_btn = filter_td.query_selector(
                "button.rgFilterButton, a.rgFilterButton, "
                "input[type='button'][class*='Filter']"
            )
            if type_btn:
                type_btn.click(timeout=3000)
                page.wait_for_timeout(500)
                starts_with = page.query_selector(
                    'li:has-text("StartsWith"), a:has-text("StartsWith")'
                )
                if starts_with:
                    starts_with.click(timeout=3000)
                    log.debug("Filter type set to StartsWith")
        except Exception as e:
            log.debug(f"Could not set filter type (non-fatal): {e}")

        # Submit the filter by pressing Enter in the input
        target.press("Enter")
        _wait_for_grid(page, timeout_ms)
        log.info(f"Grid name filter applied: starts with '{prefix}'")
        screenshot(page, "08_name_filter_applied", shots_dir)

    except Exception as e:
        log.warning(f"apply_name_filter failed ({e}) — will skip non-matching rows in loop")


def get_grid_rows(page) -> list:
    rows = page.query_selector_all("tr.rgRow, tr.rgAltRow")
    if not rows:
        rows = page.query_selector_all("tbody tr:has(td)")
    return rows


def get_row_doc_name(row) -> str:
    """
    Extract document name from a grid row.
    QT9's first two TDs are hidden (display:none) — one holds status spans
    (True/False/Active), the other holds the numeric doc ID.  The third TD
    (first visible one) is the DocumentName column.
    """
    try:
        cells = row.query_selector_all("td")
        for cell in cells:
            style = cell.get_attribute("style") or ""
            if "display:none" in style.replace(" ", ""):
                continue
            text = cell.inner_text().strip()
            if text and text != "\xa0" and len(text) > 1:
                return text
    except Exception:
        pass
    return "Unknown"


def _wait_for_grid(page, timeout_ms: int):
    """Wait for at least one grid data row after an AJAX filter/page change."""
    try:
        page.wait_for_selector(
            "tr.rgRow, tr.rgAltRow",
            state="attached",
            timeout=timeout_ms,
        )
    except PlaywrightTimeout:
        log.debug("_wait_for_grid: timed out waiting for rows — continuing anyway")


# Magic byte signatures for common office/PDF formats
_MAGIC_BYTES: dict[str, bytes] = {
    ".pdf":  b"%PDF",
    ".docx": b"PK\x03\x04",
    ".xlsx": b"PK\x03\x04",
    ".pptx": b"PK\x03\x04",
    ".doc":  b"\xd0\xcf\x11\xe0",
    ".xls":  b"\xd0\xcf\x11\xe0",
    ".ppt":  b"\xd0\xcf\x11\xe0",
}


def spot_check_downloads(output_dir: Path, sample_size: int = 10) -> bool:
    """
    Randomly sample up to `sample_size` downloaded files and verify that
    each file's leading bytes match the expected magic bytes for its extension.
    Returns True if all sampled files pass, False if any fail.
    """
    all_files = [
        f for f in output_dir.iterdir()
        if f.is_file() and f.suffix.lower() in _MAGIC_BYTES
    ]
    if not all_files:
        log.info("spot_check: no recognisable files found — skipping")
        return True

    sample = random.sample(all_files, min(sample_size, len(all_files)))
    log.info(f"spot_check: sampling {len(sample)} of {len(all_files)} file(s)")

    passed = True
    for path in sample:
        ext = path.suffix.lower()
        expected = _MAGIC_BYTES[ext]
        try:
            header = path.read_bytes()[:len(expected)]
            ok = header == expected
        except Exception as e:
            log.warning(f"  FAIL  {path.name} — read error: {e}")
            passed = False
            continue
        log.info(f"  {'PASS' if ok else 'FAIL'}  {path.name}")
        if not ok:
            log.warning(f"         expected {expected!r}, got {header!r}")
            passed = False

    if passed:
        log.info("spot_check: all sampled files passed")
    else:
        log.warning("spot_check: one or more files FAILED — check logs above")
    return passed


def next_page(page, timeout_ms: int) -> bool:
    for sel in [
        ".rgPageNext:not(.rgPagerButton[disabled])",
        "a.rgPageNext",
        "input.rgPageNext",
        "a[title='Next Page']",
    ]:
        try:
            btn = page.query_selector(sel)
            if btn:
                disabled = btn.get_attribute("disabled") or ""
                class_val = btn.get_attribute("class") or ""
                if "disabled" in disabled.lower() or "disabled" in class_val.lower():
                    return False

                # Snapshot the first row's text so we can detect an actual page change.
                # Telerik marks the last-page button disabled via CSS class rather than
                # the HTML disabled attribute, so the attribute check above can miss it.
                # If the grid doesn't change after the click we treat this as the last page.
                current_rows = page.query_selector_all("tr.rgRow, tr.rgAltRow")
                first_row_text = current_rows[0].inner_text().strip()[:120] if current_rows else ""

                btn.click(timeout=5000)

                if first_row_text:
                    try:
                        page.wait_for_function(
                            "(text) => { "
                            "const r = document.querySelectorAll('tr.rgRow, tr.rgAltRow'); "
                            "return r.length > 0 && r[0].innerText.trim().slice(0, 120) !== text; "
                            "}",
                            arg=first_row_text,
                            timeout=timeout_ms,
                        )
                    except PlaywrightTimeout:
                        log.debug("next_page: first row unchanged after click — treating as last page")
                        return False
                else:
                    _wait_for_grid(page, timeout_ms)

                log.debug("Navigated to next grid page")
                return True
        except Exception:
            continue
    return False


def download_row(page, row, doc_name: str, output_dir: Path,
                 timeout_ms: int, shots_dir: Path) -> bool:
    safe_name = sanitize_filename(doc_name)

    existing = list(output_dir.glob(f"{safe_name}.*"))
    if existing:
        log.info(f"SKIP (exists): {safe_name}")
        return True

    try:
        log.debug(f"Right-clicking row: {doc_name}")
        row.click(button="right", timeout=5000)

        # Wait for the Telerik RadContextMenu detached container to become visible
        try:
            page.wait_for_selector(
                '#ctl00_cphCenter_rcmCurrentDocsGridRow_detached',
                state='visible',
                timeout=5000
            )
        except PlaywrightTimeout:
            log.warning(f"TIMEOUT waiting for context menu: {doc_name}")
            screenshot(page, f"err_timeout_{safe_name[:40]}", shots_dir)
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
            return False

        global _ctx_screenshot_count
        log.debug("Context menu visible")
        if _ctx_screenshot_count < _CTX_SCREENSHOT_LIMIT:
            screenshot(page, f"ctx_menu_{safe_name[:40]}", shots_dir)
            _ctx_screenshot_count += 1

        # Check if "Download File" item is visible — it is hidden by QT9's JS when
        # the document has Electronic==True (no file stored in QT9 to download).
        dl_link = page.query_selector(
            '#ctl00_cphCenter_rcmCurrentDocsGridRow_detached '
            'a.rmLink:has-text("Download File")'
        )
        if not dl_link or not dl_link.is_visible():
            log.info(f"SKIP (no downloadable file — Electronic document): {doc_name}")
            page.keyboard.press("Escape")
            return True  # not a failure

        with page.expect_download(timeout=timeout_ms) as dl_info:
            dl_link.click(timeout=5000)

        download = dl_info.value
        suggested = download.suggested_filename or f"{safe_name}.bin"
        ext = Path(suggested).suffix

        filepath = output_dir / f"{safe_name}{ext}"
        counter = 1
        while filepath.exists():
            filepath = output_dir / f"{safe_name}_{counter}{ext}"
            counter += 1

        download.save_as(str(filepath))
        log.info(f"DOWNLOADED: {filepath.name}")
        return True

    except PlaywrightTimeout:
        log.warning(f"TIMEOUT waiting for download: {doc_name}")
        screenshot(page, f"err_timeout_{safe_name[:40]}", shots_dir)
        try:
            page.keyboard.press("Escape")
        except Exception:
            pass
        return False
    except Exception as e:
        log.error(f"ERROR downloading '{doc_name}': {e}")
        screenshot(page, f"err_{safe_name[:40]}", shots_dir)
        try:
            page.keyboard.press("Escape")
        except Exception:
            pass
        return False


def get_row_doc_id(row) -> str:
    """Read the numeric doc ID from TD[1] (display:none hidden cell)."""
    try:
        cells = row.query_selector_all("td")
        if len(cells) > 1:
            return cells[1].inner_text().strip()
    except Exception:
        pass
    return ""


def probe_download_url(page, row, timeout_ms: int) -> str | None:
    """
    Open the context menu on `row` and return the download URL string.
    Case A: href is a real URL — return it directly.
    Case B: href is JS or missing — trigger one real download, read its URL, cancel it.
    Returns None if "Download File" is not visible (electronic doc) or on error.
    """
    try:
        row.click(button="right", timeout=5000)
        page.wait_for_selector(
            '#ctl00_cphCenter_rcmCurrentDocsGridRow_detached',
            state='visible',
            timeout=5000,
        )
        dl_link = page.query_selector(
            '#ctl00_cphCenter_rcmCurrentDocsGridRow_detached '
            'a.rmLink:has-text("Download File")'
        )
        if not dl_link or not dl_link.is_visible():
            page.keyboard.press("Escape")
            return None

        # Case A: direct href on the anchor
        href = dl_link.get_attribute("href") or ""
        if href and not href.startswith("javascript"):
            page.keyboard.press("Escape")
            return href

        # Case B: JS-triggered — fire one real download just to capture its URL
        with page.expect_download(timeout=timeout_ms) as dl_info:
            dl_link.click(timeout=5000)
        download = dl_info.value
        url = download.url
        download.cancel()
        return url

    except Exception as e:
        log.debug(f"probe_download_url: {e}")
        try:
            page.keyboard.press("Escape")
        except Exception:
            pass
        return None


def _derive_url_template(raw_url: str, doc_id: str) -> str | None:
    """
    Replace the numeric doc ID value in `raw_url` with a `{doc_id}` placeholder.
    Returns None if the ID cannot be located in the URL.
    """
    if not doc_id or not raw_url:
        return None
    escaped = re.escape(doc_id)
    # Try DocID= parameter first (most likely)
    template = re.sub(rf'(DocID=){escaped}', r'\1{doc_id}', raw_url, flags=re.IGNORECASE)
    if '{doc_id}' not in template:
        # Generic: any query-string value equal to the numeric ID
        template = re.sub(rf'(=){escaped}(\b|&|$)', r'\1{doc_id}\2', raw_url)
    if '{doc_id}' not in template:
        log.warning(f"Could not find doc_id '{doc_id}' in URL '{raw_url}'")
        return None
    return template


def _ext_from_response(resp) -> str:
    """Derive file extension from Content-Disposition or Content-Type header."""
    cd = resp.headers.get("Content-Disposition", "")
    m = re.search(r'filename=["\']?([^"\';\s]+)', cd, re.IGNORECASE)
    if m:
        ext = Path(m.group(1).strip()).suffix
        if ext:
            return ext

    ct = resp.headers.get("Content-Type", "").split(";")[0].strip()
    if ct == "application/pdf":
        return ".pdf"
    ext = mimetypes.guess_extension(ct) or ""
    return ext


def scan_all_documents(page, prefixes: list, output_dir: Path,
                       timeout_ms: int, shots_dir: Path) -> tuple[list[dict], str | None]:
    """
    Scan all grid pages and collect document metadata without downloading anything.
    Probes the download URL template from the first eligible row that isn't already on disk.
    Returns (docs_list, url_template).
    Each doc dict: {doc_name, doc_id, exists, existing_path}.
    """
    docs = []
    url_template = None
    page_num = 1

    while True:
        rows = get_grid_rows(page)
        log.info(f"--- Scan page {page_num}: {len(rows)} rows ---")
        if page_num == 1:
            screenshot(page, f"scan_page_{page_num:03d}", shots_dir)
        if not rows:
            break

        for row in rows:
            doc_name = get_row_doc_name(row)

            if prefixes and not any(doc_name.startswith(p) for p in prefixes):
                continue

            doc_id = get_row_doc_id(row)
            safe_name = sanitize_filename(doc_name)
            existing = list(output_dir.glob(f"{safe_name}.*"))

            # Probe URL from first eligible row that hasn't been downloaded yet
            if url_template is None and not existing and doc_id:
                raw = probe_download_url(page, row, timeout_ms)
                if raw:
                    url_template = _derive_url_template(raw, doc_id)
                    if url_template:
                        log.info(f"Download URL template: {url_template}")
                    else:
                        log.warning(f"Could not derive URL template from: {raw}")
                else:
                    log.debug(f"probe_download_url returned None for '{doc_name}' (electronic doc?)")

            docs.append({
                "doc_name": doc_name,
                "doc_id": doc_id,
                "exists": bool(existing),
                "existing_path": existing[0] if existing else None,
            })

        if not next_page(page, timeout_ms):
            break
        page_num += 1

    return docs, url_template


def build_session(context) -> requests.Session:
    """Seed a requests.Session with the Playwright browser context's cookies."""
    session = requests.Session()
    for c in context.cookies():
        session.cookies.set(
            c["name"], c["value"],
            domain=c.get("domain", ""),
            path=c.get("path", "/"),
        )
    return session


def download_via_http(session, url: str, doc_name: str,
                      output_dir: Path) -> tuple[Path | None, str]:
    """
    Download one document via HTTP (no browser).
    Returns (Path, reason): reason is "ok" | "exists" | "no_file" | "error".
    """
    safe_name = sanitize_filename(doc_name)
    existing = list(output_dir.glob(f"{safe_name}.*"))
    if existing:
        return existing[0], "exists"

    try:
        resp = session.get(url, stream=True, timeout=60)
        if resp.status_code == 404:
            return None, "no_file"
        resp.raise_for_status()

        ext = _ext_from_response(resp) or ".bin"
        filepath = output_dir / f"{safe_name}{ext}"
        counter = 1
        while filepath.exists():
            filepath = output_dir / f"{safe_name}_{counter}{ext}"
            counter += 1

        with open(filepath, "wb") as fh:
            for chunk in resp.iter_content(chunk_size=65536):
                fh.write(chunk)

        return filepath, "ok"

    except Exception as e:
        log.warning(f"HTTP download failed for '{doc_name}': {e}")
        return None, "error"


def _run_sequential_fallback(base_url: str, docs_url: str, username: str, password: str,
                              prefixes: list, output_dir: Path, shots_dir: Path,
                              timeout_ms: int, args) -> None:
    """
    Sequential Playwright fallback used when the HTTP URL template cannot be determined.
    Mirrors the original main() download loop.
    """
    log.info("Running sequential Playwright fallback…")
    total_success = total_fail = 0

    with tempfile.TemporaryDirectory(prefix="qt9_tmp_") as tmp_dir, \
            sync_playwright() as p:
        browser = p.chromium.launch(
            headless=not args.headed,
            downloads_path=tmp_dir,
        )
        context = browser.new_context(
            accept_downloads=True,
            viewport={"width": 1400, "height": 900},
        )
        page = context.new_page()

        if not login(page, base_url, username, password, timeout_ms, shots_dir):
            log.error("Login failed in fallback — aborting.")
            browser.close()
            return

        page.goto(docs_url, timeout=timeout_ms)
        page.wait_for_load_state("networkidle", timeout=timeout_ms)

        apply_status_filter(page, args.status_filter, timeout_ms, shots_dir)
        set_max_page_size(page, timeout_ms, shots_dir)
        apply_name_filter(page, prefixes, timeout_ms, shots_dir)

        page_num = 1
        while True:
            rows = get_grid_rows(page)
            log.info(f"--- Grid page {page_num}: {len(rows)} rows ---")
            if not rows:
                break

            i = 0
            while i < len(rows):
                doc_name = get_row_doc_name(rows[i])
                if prefixes and not any(doc_name.startswith(p) for p in prefixes):
                    i += 1
                    continue
                if download_row(page, rows[i], doc_name, output_dir, timeout_ms, shots_dir):
                    total_success += 1
                else:
                    total_fail += 1
                    rows = get_grid_rows(page)
                    if i >= len(rows):
                        break
                    continue
                i += 1

            if not next_page(page, timeout_ms):
                break
            page_num += 1

        browser.close()

    log.info(f"Fallback done — Downloaded: {total_success}  |  Failed: {total_fail}")
    spot_check_downloads(output_dir)


def main():
    args = parse_args()
    base_url = args.url.rstrip("/")
    output_dir = Path(args.output)
    timeout_ms = args.timeout * 1000
    docs_url = base_url + DOCS_PAGE

    shots_dir = output_dir / "screenshots"
    log_dir = output_dir / "logs"

    log_file = setup_logging(log_dir)

    username, password = prompt_credentials()

    prefixes = [p.strip() for p in args.name_prefix.split(",") if p.strip()]

    log.info("=" * 60)
    log.info("QT9 QMS Bulk Document Downloader")
    log.info(f"Source  : {docs_url}")
    log.info(f"Output  : {output_dir.resolve()}")
    log.info(f"Filter  : {args.status_filter}")
    log.info(f"Prefixes: {', '.join(prefixes) if prefixes else '(none — all docs)'}")
    log.info(f"Workers : {args.workers}")
    log.info(f"Log     : {log_file}")
    log.info(f"Shots   : {shots_dir}")
    log.info("=" * 60)

    output_dir.mkdir(parents=True, exist_ok=True)
    shots_dir.mkdir(parents=True, exist_ok=True)

    # ------------------------------------------------------------------
    # Phase 1 — Playwright: login, scan all pages, probe download URL
    # ------------------------------------------------------------------
    with tempfile.TemporaryDirectory(prefix="qt9_tmp_") as tmp_dir, \
            sync_playwright() as p:
        browser = p.chromium.launch(
            headless=not args.headed,
            downloads_path=tmp_dir,
        )
        context = browser.new_context(
            accept_downloads=True,
            viewport={"width": 1400, "height": 900},
        )
        page = context.new_page()

        if not login(page, base_url, username, password, timeout_ms, shots_dir):
            log.error("Login failed — aborting.")
            browser.close()
            sys.exit(1)

        log.info(f"Navigating to documents page: {docs_url}")
        page.goto(docs_url, timeout=timeout_ms)
        page.wait_for_load_state("networkidle", timeout=timeout_ms)
        screenshot(page, "05_documents_page", shots_dir)

        apply_status_filter(page, args.status_filter, timeout_ms, shots_dir)
        set_max_page_size(page, timeout_ms, shots_dir)
        apply_name_filter(page, prefixes, timeout_ms, shots_dir)

        log.info("Phase 1: scanning all documents…")
        docs, url_template = scan_all_documents(
            page, prefixes, output_dir, timeout_ms, shots_dir
        )
        auth_session = build_session(context)
        browser.close()

    log.info(f"Phase 1 complete: {len(docs)} document(s) found")

    # ------------------------------------------------------------------
    # Phase 2 — HTTP: parallel downloads (no browser)
    # ------------------------------------------------------------------
    if not url_template:
        log.warning("Could not determine download URL template — falling back to sequential Playwright mode")
        _run_sequential_fallback(
            base_url, docs_url, username, password,
            prefixes, output_dir, shots_dir, timeout_ms, args,
        )
        return

    to_download = [d for d in docs if not d["exists"] and d["doc_id"]]
    total_skip = sum(1 for d in docs if d["exists"])
    log.info(f"Phase 2: {len(to_download)} to download, {total_skip} already on disk")

    total_success = total_fail = 0

    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        future_to_doc = {
            pool.submit(
                download_via_http,
                auth_session,
                url_template.format(doc_id=d["doc_id"]),
                d["doc_name"],
                output_dir,
            ): d
            for d in to_download
        }

        for future in as_completed(future_to_doc):
            doc = future_to_doc[future]
            try:
                path, reason = future.result()
            except Exception as exc:
                log.error(f"Unexpected error for '{doc['doc_name']}': {exc}")
                total_fail += 1
                continue

            if reason in ("ok", "exists"):
                total_success += 1
                log.info(f"DOWNLOADED: {Path(path).name}")
            elif reason == "no_file":
                total_success += 1
                log.info(f"SKIP (no file): {doc['doc_name']}")
            else:
                total_fail += 1
                log.warning(f"FAILED: {doc['doc_name']}")

    log.info("=" * 60)
    log.info(f"DONE — Downloaded: {total_success}  |  Failed: {total_fail}  |  Skipped: {total_skip}")
    log.info(f"Files : {output_dir.resolve()}")
    log.info(f"Log   : {log_file}")
    log.info("=" * 60)

    spot_check_downloads(output_dir)


if __name__ == "__main__":
    main()
