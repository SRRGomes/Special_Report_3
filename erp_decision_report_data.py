"""
SmartOps ERP decision-report downloader (Selenium/Firefox) — Credit Score Decision
Report and Risk-Based Pricing Security Approved in one script.

    python erp_decision_report_data.py --report credit-score
    python erp_decision_report_data.py --report rbp

Saves to BASE_DIR/data/<SUBDIR>/<PREFIX>_<YYYY-MM-DD>.xlsx.

Date range: the reporting period (report_period, REPORT_MODE env var) widened
back by DECISION_LOOKBACK_MONTHS whole months, because a credit decision is
usually made weeks before the facility is approved.
The ERP accepts at most ERP_MAX_RANGE_DAYS (90) days per export, so longer ranges are
downloaded in pieces (kept in <SUBDIR>/parts/) and combined into one file.
Exit code: 0 on success, 1 on any failure (so the orchestrator can stop).
"""
import os
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
DECISION_LOOKBACK_MONTHS = int(os.environ.get('DECISION_LOOKBACK_MONTHS', '2'))
# Longest From-To range the ERP report screen accepts (days, both ends inclusive)
ERP_MAX_RANGE_DAYS = int(os.environ.get('ERP_MAX_RANGE_DAYS', '90'))

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


def switch_to_iframe_with(driver, element_ids):
    """Switch into the first iframe containing all element_ids (or stay on the main page)."""
    time.sleep(2)
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    log.info(f"Found {len(iframes)} iframe(s)")
    for index, iframe in enumerate(iframes):
        try:
            driver.switch_to.frame(iframe)
            for el in element_ids:
                driver.find_element(By.ID, el)
            log.info(f"Found {element_ids} inside iframe index {index}")
            return True
        except NoSuchElementException:
            driver.switch_to.default_content()
    try:
        for el in element_ids:
            driver.find_element(By.ID, el)
        log.info(f"Found {element_ids} on main page")
        return True
    except NoSuchElementException:
        log.error(f"Could not find {element_ids} in any iframe or main page.")
        return False


def select_dropdown(driver, element_id, visible_text):
    try:
        Select(driver.find_element(By.ID, element_id)).select_by_visible_text(visible_text)
        log.info(f"Selected '{visible_text}' in '{element_id}'.")
    except Exception as e:
        log.error(f"Dropdown selection failed ({element_id}): {e}")
        raise


def get_date_range():
    """Reporting period widened back to the 1st of the month DECISION_LOOKBACK_MONTHS earlier."""
    start, end = get_period(get_mode())
    start = (start - relativedelta(months=DECISION_LOOKBACK_MONTHS)).replace(day=1)
    return start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")


def date_chunks(start, end, max_days=ERP_MAX_RANGE_DAYS):
    """Split start..end (dates, inclusive) into consecutive pieces of at most max_days days."""
    pieces = []
    while start <= end:
        stop = min(start + timedelta(days=max_days - 1), end)
        pieces.append((start, stop))
        start = stop + timedelta(days=1)
    return pieces


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
    raise TimeoutError(f"Download did not complete within {timeout}s "
                       f"(check {download_dir} for a .part file).")


def run(report_key):
    cfg = REPORTS[report_key]
    data_dir = BASE_DIR / "data" / cfg["subdir"]
    data_dir.mkdir(parents=True, exist_ok=True)

    missing = [k for k, v in (("ERP_URL", ERP_URL), ("ERP_USERNAME", ERP_USERNAME),
                              ("ERP_PASSWORD", ERP_PASSWORD)) if not v]
    if missing:
        raise EnvironmentError(f"Missing in credentials.env: {', '.join(missing)}")

    from_date, to_date = get_date_range()
    start = datetime.strptime(from_date, "%Y-%m-%d").date()
    end = datetime.strptime(to_date, "%Y-%m-%d").date()
    pieces = date_chunks(start, end)
    log.info(f"{cfg['report_name']}: {from_date} to {to_date} "
             f"(mode {get_mode()}, +{DECISION_LOOKBACK_MONTHS} month(s) look-back) | "
             f"{len(pieces)} piece(s) of at most {ERP_MAX_RANGE_DAYS} days: "
             + ", ".join(f"{a} to {b}" for a, b in pieces))

    parts_dir = data_dir / "parts"
    parts_dir.mkdir(exist_ok=True)
    today = datetime.now().strftime("%Y-%m-%d")
    t0 = time.time()
    log.info(f"Starting Firefox (download folder: {data_dir})")
    driver = setup_firefox_driver(data_dir)
    try:
        login(driver, ERP_USERNAME, ERP_PASSWORD)
        dismiss_firefox_update_dialog(UPDATE_IMG, DISMISS_IMG, timeout=15, confidence=0.8)

        parts = []
        for i, (a, b) in enumerate(pieces, start=1):
            log.info(f"Piece {i} of {len(pieces)}: {a} to {b}")
            open_report_screen(driver, cfg["report_name"], refresh=i > 1)
            set_date(driver, "search_from_date", a.strftime("%Y-%m-%d"))
            set_date(driver, "search_to_date", b.strftime("%Y-%m-%d"))

            before_files = set(os.listdir(data_dir))
            export_to_excel(driver)
            log.info("Waiting for the export to finish downloading (timeout 600s)")
            downloaded_file = Path(wait_for_download(data_dir, before_files, timeout=600))
            part = parts_dir / f"{cfg['prefix']}_{today}_part{i}_{a}_to_{b}{downloaded_file.suffix}"
            downloaded_file.replace(part)
            log.info(f"Piece {i} saved: {part.name} ({part.stat().st_size:,} bytes)")
            parts.append(part)
            time.sleep(2)

        renamed_path = data_dir / f"{cfg['prefix']}_{today}.xlsx"
        if len(parts) == 1:
            parts[0].replace(renamed_path)
        else:
            combine_parts(parts, renamed_path)
        log.info(f"{cfg['report_name']} saved: {renamed_path} "
                 f"({renamed_path.stat().st_size:,} bytes, {time.time() - t0:.0f}s)")
        return renamed_path
    finally:
        driver.quit()
        log.info("Firefox closed")


def main():
    ap = argparse.ArgumentParser(description="Download a SmartOps ERP decision report")
    ap.add_argument("--report", required=True, choices=sorted(REPORTS))
    args = ap.parse_args()
    try:
        run(args.report)
    except Exception as e:
        log.error(f"{REPORTS[args.report]['report_name']} download failed: {e}", exc_info=True)
        sys.exit(1)
    sys.exit(0)


if __name__ == "__main__":
    main()
