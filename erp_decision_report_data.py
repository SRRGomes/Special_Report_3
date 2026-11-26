"""
SmartOps ERP decision-report downloader (Selenium/Firefox) — Credit Score Decision
Report and Risk-Based Pricing Security Approved in one script.

    python erp_decision_report_data.py --report credit-score
    python erp_decision_report_data.py --report rbp
    python erp_decision_report_data.py --report rbp --lookback-months 5

Saves to BASE_DIR/data/<SUBDIR>/<PREFIX>_<YYYY-MM-DD>.xlsx.

Date range: the reporting period (report_period, REPORT_MODE env var) widened
back by the decision look-back (default 3 months; --lookback-months or
DECISION_LOOKBACK_MONTHS in credentials.env), because a credit decision is usually
made weeks before the facility is approved. The number of pieces follows from the
resulting range.
The ERP accepts at most ERP_MAX_RANGE_DAYS (90) days per export, From and To both
counted, but its PHP export also stops after 360 s ("Maximum execution time ...
exceeded") and then returns an HTML error page saved as .xlsx. So ranges are
downloaded in pieces of ERP_PIECE_DAYS (default 15, never above ERP_MAX_RANGE_DAYS),
kept in <SUBDIR>/parts/ and combined into one file. Each piece is checked as soon as
it downloads: a piece that timed out on the server is split in half and downloaded
again (down to ERP_MIN_PIECE_DAYS); a piece that returns nothing within
ERP_DOWNLOAD_TIMEOUT restarts Firefox (cancelling the stuck download) and is split the
same way; any other error page is retried, then stops the run and logs its text.
Pieces after the first reuse the open report screen (only the dates change).
Exit code: 0 on success, 1 on any failure (so the orchestrator can stop).
"""
import os
import re
import sys
import time
import argparse
from pathlib import Path
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from dotenv import load_dotenv
from igonore_firefox_updates import dismiss_firefox_update_dialog, find_firefox
from report_period import get_mode, get_period
from log_config import setup_logging

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / "credentials.env")

# Console (UTF-8) + daily log file in logs/
log = setup_logging("erp_decision_report_data")

ERP_USERNAME = os.environ.get('ERP_USERNAME')
ERP_PASSWORD = os.environ.get('ERP_PASSWORD')
ERP_URL = os.environ.get('ERP_URL')

# Months of decisions to pull before the reporting period start (override in credentials.env)
DEFAULT_DECISION_LOOKBACK_MONTHS = 3
DECISION_LOOKBACK_MONTHS = int(os.environ.get('DECISION_LOOKBACK_MONTHS',
                                              str(DEFAULT_DECISION_LOOKBACK_MONTHS)))
# Longest From-To range the ERP report screen accepts (days, both ends inclusive)
ERP_MAX_RANGE_DAYS = int(os.environ.get('ERP_MAX_RANGE_DAYS', '90'))
# Days per download piece: small enough for the ERP to build the .xlsx within its
# 360 s PHP limit. Pieces that still time out are halved, down to ERP_MIN_PIECE_DAYS.
ERP_PIECE_DAYS = min(int(os.environ.get('ERP_PIECE_DAYS', '15')), ERP_MAX_RANGE_DAYS)
ERP_MIN_PIECE_DAYS = max(int(os.environ.get('ERP_MIN_PIECE_DAYS', '1')), 1)
# Retries for a piece that comes back as an error page other than the timeout
ERP_PIECE_RETRIES = int(os.environ.get('ERP_PIECE_RETRIES', '1'))
# Seconds to wait for a piece's file to appear; a piece with no file by then is treated
# like a server timeout (Firefox restarted, piece split in half)
ERP_DOWNLOAD_TIMEOUT = int(os.environ.get('ERP_DOWNLOAD_TIMEOUT', '600'))
# Seconds to wait for the report screen's dropdowns to load
ERP_SCREEN_TIMEOUT = int(os.environ.get('ERP_SCREEN_TIMEOUT', '30'))

REPORTS = {
    "credit-score": {
        "report_name": "Credit Score Decision Report",
        "subdir": "CREDIT_SCORE_DECISION_data",
        "prefix": "CREDIT_SCORE_DECISION",
    },
    "rbp": {
        "report_name": "Risk-Based Pricing Security Approved",
        "subdir": "RBP_SECURITY_APPROVED_data",
        "prefix": "RBP_SECURITY_APPROVED",
    },
}

UPDATE_IMG = str(BASE_DIR / "img" / "Update_Message.png")
DISMISS_IMG = str(BASE_DIR / "img" / "Dismiss_Button.png")


def setup_firefox_driver(download_dir):
    options = Options()
    firefox = find_firefox()          # same lookup as the Oracle BI downloader
    if firefox:
        options.binary_location = firefox
        log.info(f"Firefox binary: {firefox}")
    options.set_preference("webdriver_accept_untrusted_certs", True)
    options.set_preference("acceptInsecureCerts", True)
    options.set_preference("browser.download.folderList", 2)
    options.set_preference("browser.download.dir", str(download_dir))
    options.set_preference("browser.helperApps.neverAsk.saveToDisk",
        "application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream")
    options.set_preference("pdfjs.disabled", True)
    # options.headless = True
    driver = webdriver.Firefox(options=options)
    driver.maximize_window()
    return driver


def login(driver, username, password):
    log.info("Navigating to login page")
    driver.get(ERP_URL)
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "txtusername")))
        driver.find_element(By.ID, "txtusername").send_keys(username)
        driver.find_element(By.ID, "txtpassword").send_keys(password)
        driver.find_element(By.CLASS_NAME, "login100-form-btn").click()
        log.info(f"Login submitted for user {username}")
    except Exception as e:
        log.error(f"Login failed: {e}")
        raise


def expand_menu(driver, menu_text):
    log.info(f"Expanding menu: {menu_text}")
    dropdown_buttons = driver.find_elements(By.CLASS_NAME, "dropdown-btn")
    for btn in dropdown_buttons:
        try:
            span = btn.find_element(By.TAG_NAME, "span")
            if span.text.strip() == menu_text:
                driver.execute_script("arguments[0].click();", btn)
                time.sleep(1)
                return True
        except Exception:
            continue
    log.warning(f"Menu '{menu_text}' not found.")
    return False


def click_reports(driver):
    log.info("Clicking 'Reporting Management' link")
    try:
        reporting_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Reporting Management')]"))
        )
        driver.execute_script("arguments[0].click();", reporting_link)
        log.info("Navigated to Reports.")
    except TimeoutException:
        log.error("'Reporting Management' link not found or not clickable.")
        raise


def _find_in_frames(driver, element_ids):
    """One pass: switch into the first iframe (or the main page) holding all element_ids."""
    driver.switch_to.default_content()
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    for index, iframe in enumerate(iframes):
        try:
            driver.switch_to.frame(iframe)
            for el in element_ids:
                driver.find_element(By.ID, el)
            return f"iframe index {index} (of {len(iframes)})"
        except Exception:
            driver.switch_to.default_content()
    try:
        for el in element_ids:
            driver.find_element(By.ID, el)
        return f"main page ({len(iframes)} iframe(s))"
    except NoSuchElementException:
        return None


def switch_to_iframe_with(driver, element_ids, timeout=None):
    """Poll until element_ids are found in an iframe or the main page (slow ERP pages)."""
    timeout = ERP_SCREEN_TIMEOUT if timeout is None else timeout
    deadline = time.time() + timeout
    time.sleep(1)
    while True:
        where = _find_in_frames(driver, element_ids)
        if where:
            log.info(f"Found {element_ids} in {where}")
            return True
        if time.time() >= deadline:
            log.error(f"Could not find {element_ids} in any iframe or main page within {timeout}s.")
            return False
        time.sleep(1)


def select_dropdown(driver, element_id, visible_text):
    try:
        Select(driver.find_element(By.ID, element_id)).select_by_visible_text(visible_text)
        log.info(f"Selected '{visible_text}' in '{element_id}'.")
    except Exception as e:
        log.error(f"Dropdown selection failed ({element_id}): {e}")
        raise


def get_date_range(lookback_months=None):
    """Reporting period widened back to the 1st of the month `lookback_months` earlier."""
    lookback = DECISION_LOOKBACK_MONTHS if lookback_months is None else lookback_months
    start, end = get_period(get_mode())
    start = (start - relativedelta(months=lookback)).replace(day=1)
    return start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")


def date_chunks(start, end, max_days=ERP_PIECE_DAYS):
    """Split start..end (dates, inclusive) into consecutive pieces of at most max_days days."""
    pieces = []
    while start <= end:
        stop = min(start + timedelta(days=max_days - 1), end)
        pieces.append((start, stop))
        start = stop + timedelta(days=1)
    return pieces


class ERPExportError(RuntimeError):
    """The ERP returned an error page (saved as .xlsx) instead of a workbook."""


class ERPServerTimeout(ERPExportError):
    """The ERP's PHP export hit its maximum execution time; a smaller range may work."""


class DownloadTimeout(TimeoutError):
    """No file arrived for a piece within the wait time."""


XLSX_SIGNATURE = b"PK\x03\x04"      # .xlsx is a zip archive
TIMEOUT_PATTERNS = ("maximum execution time", "max_execution_time")


def error_page_text(raw_bytes, limit=300):
    """Readable text of an HTML/PHP error page: tags removed, whitespace collapsed."""
    text = raw_bytes.decode("utf-8", "replace")
    text = re.sub(r"<[^>]+>", " ", text)
    return re.sub(r"\s+", " ", text).strip()[:limit]


def check_piece(path):
    """Raise ERPServerTimeout / ERPExportError if the downloaded piece is not a real .xlsx."""
    raw_bytes = Path(path).read_bytes()
    if raw_bytes[:4] == XLSX_SIGNATURE:
        return
    text = error_page_text(raw_bytes)
    if any(p in text.lower() for p in TIMEOUT_PATTERNS):
        raise ERPServerTimeout(text)
    raise ERPExportError(text or f"{len(raw_bytes)} bytes, not an Excel file")


def split_piece(a, b):
    """Two halves of a..b (inclusive)."""
    mid = a + timedelta(days=((b - a).days + 1) // 2 - 1)
    return (a, mid), (mid + timedelta(days=1), b)


def _header_index(raw):
    """Row index of the header in a raw (header=None) export: fullest of the first 30 rows."""
    counts = raw.head(30).notna().sum(axis=1)
    return int(counts.idxmax())


def combine_parts(parts, out_path):
    """Stack the data rows of several exports of the same report under one header row."""
    frames, header = [], None
    for path in parts:
        raw = pd.read_excel(path, header=None)
        h = _header_index(raw)
        cols = [str(c).strip() for c in raw.iloc[h].tolist()]
        if header is None:
            header = cols
        elif cols != header:
            raise ValueError(f"{Path(path).name}: columns differ from the first piece "
                             f"({len(cols)} vs {len(header)}); cannot combine")
        body = raw.iloc[h + 1:].copy()
        body.columns = header
        frames.append(body)
        log.info(f"Piece {Path(path).name}: {len(body):,} rows")
    combined = pd.concat(frames, ignore_index=True)
    combined.to_excel(out_path, index=False)
    log.info(f"Combined {len(parts)} pieces -> {Path(out_path).name} ({len(combined):,} rows)")
    return out_path


def open_report_screen(driver, report_name, refresh=False):
    """Menus -> Reporting Management -> Help Desk -> report. refresh=True starts from a fresh page."""
    driver.switch_to.default_content()
    if refresh:
        driver.refresh()
        log.info("Page refreshed for the next date piece")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "dropdown-btn")))
    time.sleep(2)
    expand_menu(driver, "Control Panel")
    expand_menu(driver, "Report")
    click_reports(driver)
    if not switch_to_iframe_with(driver, ["search_module", "search_report"]):
        raise RuntimeError("Could not find 'search_module' dropdown.")
    select_dropdown(driver, "search_module", "Help Desk")
    if not switch_to_iframe_with(driver, ["search_report"]):
        raise RuntimeError("Could not find 'search_report' dropdown.")
    select_dropdown(driver, "search_report", report_name)


def set_date(driver, element_id, date_str):
    try:
        el = driver.find_element(By.ID, element_id)
        driver.execute_script("arguments[0].value = arguments[1];", el, date_str)
        log.info(f"Set '{element_id}' to {date_str}")
    except Exception as e:
        log.error(f"Failed to set '{element_id}': {e}")
        raise


def export_to_excel(driver):
    try:
        export_btn = driver.find_element(By.ID, "excel_btn")

        if export_btn.get_attribute("disabled"):
            log.warning("Excel button is disabled. Waiting for it to be enabled...")
            WebDriverWait(driver, 30).until(
                lambda d: d.find_element(By.ID, "excel_btn").get_attribute("disabled") is None
            )
            log.info("Excel button is now enabled")
            export_btn = driver.find_element(By.ID, "excel_btn")

        driver.execute_script("arguments[0].click();", export_btn)
        log.info("Clicked 'Excel' export button.")

        # Confirmation modal (if any)
        try:
            WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.CLASS_NAME, "ajs-button"))
            )
            yes_button = driver.find_element(By.XPATH, "//button[contains(@class, 'ajs-ok')]")
            driver.execute_script("arguments[0].click();", yes_button)
            log.info("Clicked 'Yes' to confirm export.")
        except TimeoutException:
            log.info("No confirmation modal found - download may have started directly")
    except Exception as e:
        log.error(f"Failed to export: {e}")
        raise


def wait_for_download(download_dir, before, timeout=600, poll_interval=5):
    """
    Wait for a NEW .xlsx (not in the `before` snapshot) to finish downloading.
    Firefox writes '<name>.part' and a 0-byte '<name>.xlsx' placeholder; the file
    counts as done when it is non-empty, has no .part companion and its size is
    stable across two polls.
    """
    elapsed = 0
    last_size = -1
    while elapsed < timeout:
        after = set(os.listdir(download_dir))
        new_xlsx = sorted(f for f in (after - before) if f.endswith('.xlsx'))
        if new_xlsx:
            file_path = os.path.join(download_dir, new_xlsx[0])
            size = os.path.getsize(file_path)
            if (size > 0 and size == last_size
                    and not os.path.exists(file_path + ".part")):
                return file_path
            if size != last_size:
                log.info(f"Downloading {new_xlsx[0]}: {size:,} bytes")
            last_size = size
        time.sleep(poll_interval)
        elapsed += poll_interval
    pending = sorted(f for f in set(os.listdir(download_dir)) - before)
    raise DownloadTimeout(f"no file within {timeout}s; new items in the download folder: "
                          f"{', '.join(pending) or 'none'}")


def set_piece_dates(driver, report_name, a, b, first):
    """
    Set From/To for a piece. Pieces after the first stay on the open report screen
    and only change the dates; if that screen is gone, re-open it once and retry.
    """
    try:
        set_date(driver, "search_from_date", a.strftime("%Y-%m-%d"))
        set_date(driver, "search_to_date", b.strftime("%Y-%m-%d"))
    except Exception as e:
        if first:
            raise
        log.warning(f"Report screen not available for the next piece ({e.__class__.__name__}); "
                    f"re-opening it")
        open_report_screen(driver, report_name, refresh=True)
        set_date(driver, "search_from_date", a.strftime("%Y-%m-%d"))
        set_date(driver, "search_to_date", b.strftime("%Y-%m-%d"))


def check_credentials():
    missing = [k for k, v in (("ERP_URL", ERP_URL), ("ERP_USERNAME", ERP_USERNAME),
                              ("ERP_PASSWORD", ERP_PASSWORD)) if not v]
    if missing:
        raise EnvironmentError(f"Missing in credentials.env: {', '.join(missing)}")


def start_firefox(data_dir):
    """New Firefox, logged in to the ERP."""
    log.info(f"Starting Firefox (download folder: {data_dir})")
    driver = setup_firefox_driver(data_dir)
    login(driver, ERP_USERNAME, ERP_PASSWORD)
    dismiss_firefox_update_dialog(UPDATE_IMG, DISMISS_IMG, timeout=15, confidence=0.8)
    return driver


def open_report_screen_with_retry(driver, report_name):
    try:
        open_report_screen(driver, report_name)
    except Exception as e:
        log.warning(f"Report screen did not open ({e}); refreshing and trying once more")
        open_report_screen(driver, report_name, refresh=True)


def download_piece(driver, cfg, a, b, n, data_dir, parts_dir, today):
    """
    Export one date piece and check it. Returns the saved part file.
    Raises:
      DownloadTimeout   no file arrived (Firefox is closed here so the stuck download stops)
      ERPServerTimeout  the ERP returned its "maximum execution time" error page
      ERPExportError    the ERP returned another error page
    Error pages are kept as parts/*.error.html.
    """
    before_files = set(os.listdir(data_dir))
    t0 = time.time()
    export_to_excel(driver)
    log.info(f"Waiting for the export to finish downloading (timeout {ERP_DOWNLOAD_TIMEOUT}s)")
    try:
        downloaded_file = Path(wait_for_download(data_dir, before_files,
                                                 timeout=ERP_DOWNLOAD_TIMEOUT))
    except DownloadTimeout as e:
        shot = parts_dir / f"{cfg['prefix']}_{today}_part{n}_{a}_to_{b}_no_download.png"
        try:
            driver.save_screenshot(str(shot))
            log.warning(f"Screenshot of the ERP screen: {shot.name}")
        except Exception:
            pass
        log.warning(f"Piece {a} to {b}: {e}")
        # Close Firefox so a late-arriving file can't be taken for the next piece
        driver.quit()
        log.info("Firefox closed (stuck download cancelled)")
        for f in set(os.listdir(data_dir)) - before_files:
            try:
                (data_dir / f).unlink()
                log.info(f"Removed leftover {f}")
            except OSError:
                pass
        raise

    part = parts_dir / f"{cfg['prefix']}_{today}_part{n}_{a}_to_{b}{downloaded_file.suffix}"
    downloaded_file.replace(part)
    size = part.stat().st_size
    try:
        check_piece(part)
    except ERPExportError as e:
        bad = part.with_suffix(".error.html")
        part.replace(bad)
        e.details = f"{size:,} bytes after {time.time() - t0:.0f}s, kept as {bad.name}"
        e.bad_file = bad
        raise
    log.info(f"Piece {n} saved: {part.name} ({size:,} bytes, {time.time() - t0:.0f}s)")
    return part


def split_or_fail(queue, a, b, error):
    """Put the two halves of a..b at the front of the queue, or stop if it can't shrink."""
    days = (b - a).days + 1
    if days <= ERP_MIN_PIECE_DAYS:
        raise type(error)(f"{a} to {b} still fails at {days} day(s) "
                          f"(ERP_MIN_PIECE_DAYS={ERP_MIN_PIECE_DAYS}): {error}") from None
    halves = split_piece(a, b)
    queue[:0] = list(halves)
    log.info("Splitting into " + " and ".join(f"{x} to {y}" for x, y in halves))


def run(report_key, lookback_months=None):
    """
    1. Work out the date range and cut it into pieces (ERP_PIECE_DAYS each)
    2. Log in once, then download the pieces one by one on the same report screen:
         good file           -> keep it
         nothing arrives     -> restart Firefox, split the piece in half
         ERP timeout page    -> split the piece in half
         other error page    -> retry (ERP_PIECE_RETRIES), then stop
    3. Combine the pieces into <PREFIX>_<today>.xlsx
    """
    cfg = REPORTS[report_key]
    data_dir = BASE_DIR / "data" / cfg["subdir"]
    parts_dir = data_dir / "parts"
    parts_dir.mkdir(parents=True, exist_ok=True)
    check_credentials()

    # 1. Date range and pieces
    lookback = DECISION_LOOKBACK_MONTHS if lookback_months is None else lookback_months
    from_date, to_date = get_date_range(lookback)
    start = datetime.strptime(from_date, "%Y-%m-%d").date()
    end = datetime.strptime(to_date, "%Y-%m-%d").date()
    pieces = date_chunks(start, end)
    log.info(f"{cfg['report_name']}: {from_date} to {to_date} "
             f"(mode {get_mode()}, +{lookback} month(s) look-back) | "
             f"{len(pieces)} piece(s) of at most {ERP_PIECE_DAYS} days: "
             + ", ".join(f"{a} to {b}" for a, b in pieces))

    today = datetime.now().strftime("%Y-%m-%d")
    t0 = time.time()
    driver = None
    try:
        # 2. Download the pieces
        driver = start_firefox(data_dir)
        need_screen = True            # open the report screen before the next piece
        queue = list(pieces)          # chronological; halves of a failed piece go to the front
        parts, retries, n = [], {}, 0
        while queue:
            a, b = queue.pop(0)
            n += 1
            log.info(f"Piece {n} ({len(queue)} more queued): {a} to {b} ({(b - a).days + 1} days)")
            if need_screen:
                open_report_screen_with_retry(driver, cfg["report_name"])
            set_piece_dates(driver, cfg["report_name"], a, b, first=need_screen)
            need_screen = False

            try:
                parts.append(download_piece(driver, cfg, a, b, n, data_dir, parts_dir, today))
                time.sleep(2)
            except DownloadTimeout as e:              # nothing arrived
                driver = None                         # closed in download_piece
                split_or_fail(queue, a, b, e)
                driver = start_firefox(data_dir)
                need_screen = True
            except ERPServerTimeout as e:             # ERP gave up (360 s limit)
                log.warning(f"Piece {a} to {b}: ERP timed out ({e.details}): {str(e)[:150]}")
                split_or_fail(queue, a, b, e)
                time.sleep(2)
            except ERPExportError as e:               # any other error page
                tries = retries.get((a, b), 0)
                if tries >= ERP_PIECE_RETRIES:
                    raise ERPExportError(f"{a} to {b}: ERP returned an error page instead of "
                                         f"Excel (see {e.bad_file}): {e}") from None
                retries[(a, b)] = tries + 1
                queue.insert(0, (a, b))
                log.warning(f"Piece {a} to {b}: ERP returned an error page ({e.details}); "
                            f"retry {tries + 1} of {ERP_PIECE_RETRIES}: {e}")
                time.sleep(5)

        # 3. One file for the summary scripts
        renamed_path = data_dir / f"{cfg['prefix']}_{today}.xlsx"
        if len(parts) == 1:
            parts[0].replace(renamed_path)
        else:
            combine_parts(parts, renamed_path)
        log.info(f"{cfg['report_name']} saved: {renamed_path} "
                 f"({renamed_path.stat().st_size:,} bytes, {time.time() - t0:.0f}s)")
        return renamed_path
    finally:
        if driver is not None:
            driver.quit()
            log.info("Firefox closed")


def main():
    ap = argparse.ArgumentParser(description="Download a SmartOps ERP decision report")
    ap.add_argument("--report", required=True, choices=sorted(REPORTS))
    ap.add_argument("--lookback-months", type=int,
                    help=f"months of decisions before the period start (default "
                         f"{DECISION_LOOKBACK_MONTHS}: DECISION_LOOKBACK_MONTHS or "
                         f"{DEFAULT_DECISION_LOOKBACK_MONTHS})")
    args = ap.parse_args()
    try:
        if args.lookback_months is not None and args.lookback_months < 0:
            raise ValueError("--lookback-months must be 0 or more")
        run(args.report, args.lookback_months)
    except Exception as e:
        log.error(f"{REPORTS[args.report]['report_name']} download failed: {e}", exc_info=True)
        sys.exit(1)
    sys.exit(0)


if __name__ == "__main__":
    main()
