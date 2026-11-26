"""
Approval Report downloader (Oracle BI, pyautogui/Firefox) - HO and Branch in one script.
Used by the Approved Facility report.

HOW IT WORKS (run_process() follows these steps; each screen element is found by its
image in img/, so the screen must look like it did when the images were captured -
see check_images.py)
  1. Dates       period of REPORT_MODE (report_period.py) in Oracle BI format MM-DD-YYYY
  2. Firefox     open the Oracle BI login page, dismiss the update dialog, sign in
  3. Report      search the catalog, open the report, pick HO or BRANCHES
  4. Filters     type the From / To dates, Apply
  5. Download    export to Excel, wait for the file in Downloads
  6. Save        move it to data/APPROVAL_REPORT_data_<HO|BRANCH>/APPROVAL_REPORT_<X>_<stamp>.xlsx
  main() retries the whole process up to 3 times (screenshot + browser closed after a failure).

    python approval_report_data.py --type HO
    python approval_report_data.py --type BRANCHES
"""
import sys
import argparse
import pyautogui
import time
from datetime import datetime
from dotenv import load_dotenv
import os
import shutil
from pathlib import Path
import psutil
from report_period import get_mode, get_period
from log_config import setup_logging

# NOTE: module name on disk is misspelled ("igonore" not "ignore").
# Rename the file to ignore_firefox_updates.py if you want to fix it; until
# then this import must match the actual filename.
from igonore_firefox_updates import (
    dismiss_firefox_update_dialog,
    oracle_sign_in,
    search_bar_enter,
    open_catalog_report,
    set_date_field,
    click_button,
    click_toolbar_option,
    open_in_firefox,
)


# Console (UTF-8) + daily log file in logs/
log = setup_logging("approval_report_data")

# Anchor all paths to this script's folder so every pipeline script shares
# one layout: BASE_DIR/data/<subdir>, BASE_DIR/reports/<date>, etc.
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"          # staging; the summary moves used files to reports/<date>/data

# --type value -> folder/file suffix
TYPE_SUFFIX = {"HO": "HO", "BRANCHES": "BRANCH"}


# ----------------------------
# Utility Functions
# ----------------------------

def cleanup_orphaned_temp_files(prefix="APPROVAL_REPORT"):
    """Remove leftover .part/.crdownload temp files from crashed attempts."""
    downloads_path = Path.home() / "Downloads"
    for pattern in (f"{prefix}*.part", f"{prefix}*.crdownload"):
        for f in downloads_path.glob(pattern):
            try:
                f.unlink()
                log.info(f"Removed orphaned temp file: {f.name}")
            except OSError as e:
                log.warning(f"Could not remove {f.name}: {e}")


def wait_for_download(prefix, started_after, timeout=600):
    """Wait until a NEW file starting with prefix is fully downloaded.

    FIXED vs previous version:
      - st_birthtime removed (AttributeError on Windows Python <= 3.11);
        st_ctime IS creation time on Windows.
      - Firefox writes a 0-byte `name` placeholder next to `name.part`;
        the old name-ends-with-.part check passed on the placeholder
        immediately. Now completion requires: no .part/.crdownload
        companion, nonzero size, and size stable across two polls.
      - Only files created after `started_after` count, so stale files
        (e.g. the other report type's file from an earlier run) can't satisfy the wait.
      - Timeout 120s -> 600s so slow reports don't get killed mid-download.
    """
    downloads_path = Path.home() / "Downloads"
    deadline = time.time() + timeout
    last_size = {}  # path -> size at previous poll

    while time.time() < deadline:
        candidates = [
            f for f in downloads_path.glob(f"{prefix}*")
            if f.is_file()
            and not f.name.endswith((".part", ".crdownload"))
            and f.stat().st_ctime >= started_after
        ]

        for f in sorted(candidates, key=lambda p: p.stat().st_ctime, reverse=True):
            part_companion = f.with_name(f.name + ".part")
            cr_companion = f.with_name(f.name + ".crdownload")
            if part_companion.exists() or cr_companion.exists():
                if part_companion.exists():
                    log.info(f"Still downloading: {f.name} "
                          f"({part_companion.stat().st_size:,} bytes so far)")
                else:
                    log.info(f"Still downloading: {f.name}")
                continue

            size = f.stat().st_size
            if size == 0:
                continue  # Firefox placeholder not yet finalized

            if last_size.get(f) == size:
                # Size stable across two polls, no temp companion -> complete
                return f
            last_size[f] = size
            log.info(f"Verifying size stability: {f.name} ({size:,} bytes)")

        time.sleep(5)

    raise TimeoutError("Download did not complete within the timeout. "
                       "Check Downloads for a lingering .part file.")


def move_report_file(started_after, suffix: str):
    """Move downloaded report from Downloads to BASE_DIR/data/APPROVAL_REPORT_data_<SUFFIX>.

    Anchored to BASE_DIR so the download folder is the same one the
    summary script reads, regardless of who launches the script or from
    which working directory.
    """
    target_path = DATA_DIR / f"APPROVAL_REPORT_data_{suffix.upper()}"
    os.makedirs(target_path, exist_ok=True)

    prefix = "APPROVAL_REPORT"

    log.info("Waiting for report to finish downloading...")

    downloaded_file = wait_for_download(prefix, started_after)

    # Rename and move the file
    latest_file = Path(downloaded_file)

    # Create new filename with the suffix while preserving extension
    current_datetime = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    new_filename = f"{prefix}_{suffix.upper()}_{current_datetime}{latest_file.suffix}"
    renamed_path = latest_file.parent / new_filename

    # Rename in Downloads
    renamed_file = latest_file.rename(renamed_path)

    # Move to target location
    destination = target_path / renamed_file.name
    shutil.move(str(renamed_file), str(destination))

    # Double-check deletion (verify file no longer in Downloads, exists in target, and is not empty)
    if not renamed_path.exists() and destination.exists() and destination.stat().st_size > 0:
        log.info(f"Report ({destination.stat().st_size:,} bytes) moved to: {destination} and removed from Downloads")
    else:
        raise RuntimeError("File move verification failed — file missing or empty!")

    return destination


def _require(ok, step):
    """Raise if a UI step returned False, so the retry loop can catch it."""
    if not ok:
        raise RuntimeError(f"UI step failed: {step}")
    log.info(f"UI step OK: {step}")


def close_browser():
    """Close Firefox by process name — never Alt+F4.

    FIXED: Alt+F4 went to whichever window had focus (download dialogs,
    the terminal, etc.), so it sometimes closed the wrong window or
    killed Firefox mid-download. Terminate gracefully first so Firefox
    flushes its download state, then kill anything that ignores it.
    """
    procs = [p for p in psutil.process_iter(['name'])
             if p.info['name'] and p.info['name'].lower() == 'firefox.exe']
    if not procs:
        log.info("No Firefox process found to close.")
        return

    for p in procs:
        try:
            p.terminate()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

    gone, alive = psutil.wait_procs(procs, timeout=10)
    for p in alive:
        try:
            p.kill()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    log.info(f"Closed Firefox ({len(procs)} process(es)).")


def take_screenshot(attempt):
    """Take a screenshot on failure. Never raise — cleanup must not mask the real error."""
    try:
        screenshot_path = BASE_DIR / f"report_fail_attempt_{attempt}.png"
        pyautogui.screenshot(str(screenshot_path))
        log.info(f"Screenshot saved: {screenshot_path}")
    except Exception as e:
        log.warning(f"Could not take screenshot: {e}")


def _safe_close_browser():
    """Never let browser-close failure mask the original error."""
    try:
        close_browser()
    except Exception as e:
        log.warning(f"Error closing browser: {e}")


# ----------------------------
# Main Automation
# ----------------------------

def run_process(report_type):
    """Run the report generation process once. report_type: 'HO' or 'BRANCHES'."""
    # Load environment variables from .env file (anchored to BASE_DIR)
    load_dotenv(BASE_DIR / 'credentials.env')

    # Configuration from .env file
    LOGIN_URL = os.getenv('ORACLE_URL')
    USERNAME = os.getenv('ORACLE_USERNAME')
    PASSWORD = os.getenv('ORACLE_PASSWORD')
    SEARCH_TEXT = 'APPROVAL_REPORT'

    missing = [k for k, v in (("ORACLE_URL", LOGIN_URL),
                              ("ORACLE_USERNAME", USERNAME),
                              ("ORACLE_PASSWORD", PASSWORD)) if not v]
    if missing:
        raise EnvironmentError(
            f"Missing required env vars in credentials.env: {', '.join(missing)}"
        )
    log.info(f"Credentials loaded for user {USERNAME}")

    # Report type: "HO" or "BRANCHES" (from --type).
    REPORT_TYPE = report_type

    # ---- Image paths (anchored to BASE_DIR) ----
    oracle_sign_in_page = str(BASE_DIR / "img" / "approval" / "Oracle_Signin_Page.png")
    oracle_signed_in_page = str(BASE_DIR / "img" / "approval" / "Half_UI.png")
    search_bar_image = str(BASE_DIR / "img" / "approval" / "Search_Bar.png")
    search_result_image = str(BASE_DIR / "img" / "approval" / "Approval_Report.png")
    open_report_button_image = str(BASE_DIR / "img" / "approval" / "Open_Approval_Report.png")
    report_type_field_image = str(BASE_DIR / "img" / "approval" / "Report_Type_Field.png")
    report_type_ho_image = str(BASE_DIR / "img" / "approval" / "Report_Type_HO.png")
    report_type_branches_image = str(BASE_DIR / "img" / "approval" / "Report_Type_Branches.png")
    from_date_field_image = str(BASE_DIR / "img" / "approval" / "From_Date.png")
    to_date_field_image = str(BASE_DIR / "img" / "approval" / "To_Date.png")
    apply_button_image = str(BASE_DIR / "img" / "approval" / "Apply.png")
    download_icon_image = str(BASE_DIR / "img" / "approval" / "Download_Icon.png")
    download_excel_option_image = str(BASE_DIR / "img" / "approval" / "Excel_Option.png")

    report_type_option_image = (report_type_ho_image if REPORT_TYPE == "HO"
                                else report_type_branches_image)

    # ---- Date Calculation (mode-aware, MM-DD-YYYY) ----
    # Mode definitions live in report_period.py
    mode = get_mode()
    period_start, period_end = get_period(mode)
    from_date = period_start.strftime("%m-%d-%Y")
    to_date = period_end.strftime("%m-%d-%Y")
    log.info(f"Approval Report {REPORT_TYPE} | mode {mode} | From {from_date} To {to_date}")

    # Clean orphaned temp files, then record start time: only files
    # created after this moment can satisfy wait_for_download.
    cleanup_orphaned_temp_files()
    started_after = time.time()

    #pyautogui.press('f11')
    # Open the login URL in Firefox explicitly. No fallback to the default
    # browser: the screen images and close_browser() are Firefox-specific.
    open_in_firefox(LOGIN_URL)
    time.sleep(7)  # Wait for page to load

    # ---- Ignore Firefox Updates ----
    update_img = str(BASE_DIR / "img" / "Update_Message.png")
    dismiss_img = str(BASE_DIR / "img" / "Dismiss_Button.png")
    dismiss_firefox_update_dialog(update_img, dismiss_img, timeout=15, confidence=0.8)

    # ---- Login ----
    _require(oracle_sign_in(USERNAME, PASSWORD, signin_page_image=oracle_sign_in_page,
                            signed_in_image=oracle_signed_in_page, timeout=15, confidence=0.7),
             "Oracle sign-in")
    time.sleep(1)

    # ---- Search ----
    _require(search_bar_enter(SEARCH_TEXT, search_bar_image, timeout=15, confidence=0.7),
             "Search bar enter")
    time.sleep(0.5)

    # ---- Search Result Selection ----
    _require(open_catalog_report(search_result_image, open_report_button_image,
                                 timeout=15, confidence=0.7),
             "Open catalog report")
    time.sleep(1)

    # ---- Select Report Type ----
    _require(click_toolbar_option(report_type_field_image, report_type_option_image,
                                  timeout=15, confidence=0.7),
             f"Select report type '{REPORT_TYPE}'")
    time.sleep(1)

    # ---- Date Selection ----
    # Enter date in From date field
    _require(set_date_field(from_date_field_image, from_date, timeout=5, confidence=0.7),
             "Set from date field")
    time.sleep(1)

    # Enter date in To date field
    _require(set_date_field(to_date_field_image, to_date, timeout=5, confidence=0.7),
             "Set to date field")
    time.sleep(1)

    _require(click_button(apply_button_image, timeout=15, confidence=0.7),
             "Click apply button")
    log.info("Waiting 20s for the report to render after Apply")
    time.sleep(20)

    # ---- Generate Report ----
    _require(click_toolbar_option(download_icon_image, download_excel_option_image,
                                  timeout=15, confidence=0.7),
             "Generate and download report")
    time.sleep(1)

    msg2 = str(BASE_DIR / "img" / "Message.png")
    button2 = str(BASE_DIR / "img" / "Button.png")
    dismiss_firefox_update_dialog(msg2, button2)

    # ---- Wait & Move Report ----
    return move_report_file(started_after, TYPE_SUFFIX[REPORT_TYPE])


def main():
    ap = argparse.ArgumentParser(description="Download the Oracle BI Approval Report")
    ap.add_argument("--type", required=True, choices=sorted(TYPE_SUFFIX),
                    help="Approval Report type to download")
    report_type = ap.parse_args().type
    label = TYPE_SUFFIX[report_type]
    run_start = time.time()

    max_attempts = 3
    success = False
    for attempt in range(1, max_attempts + 1):
        try:
            log.info(f"Attempt {attempt} of {max_attempts} ({label})")
            destination = run_process(report_type)
            success = True
            log.info(f"Success: report saved at {destination} "
                     f"(attempt {attempt}, {time.time() - run_start:.0f}s)")
            break  # Exit loop on success

        except Exception as e:
            log.error(f"Attempt {attempt} failed: {e}", exc_info=True)
            take_screenshot(f"{label}_{attempt}")
            _safe_close_browser()
            if attempt < max_attempts:
                log.warning(f"Retrying in 5s (attempt {attempt + 1} of {max_attempts})")
                time.sleep(5)  # wait before retry
            else:
                log.error("All attempts failed. Exiting.")
                # FIXED: no `return` here — it skipped the final
                # close_browser() AND sys.exit(1), so the script
                # reported exit code 0 to the orchestrator on total
                # failure. Fall through to the shared exit logic.

    # Always ensure browser is closed at the end
    _safe_close_browser()
    if success:
        log.info(f"Approval Report ({label}) download process completed successfully.")
        sys.exit(0)
    else:
        log.error(f"Approval Report ({label}) download process failed after all attempts "
                  f"({time.time() - run_start:.0f}s).")
        sys.exit(1)


if __name__ == "__main__":
    main()