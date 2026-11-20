"""
BRANCH DISBURSEMENT OLD downloader (Oracle BI, pyautogui/Firefox) - customer history
for the Existing / New Customer check of the Approved Facility (AF / 3W / 2W) Analysis.

The report lists facilities DISBURSED between From and To (with NIC_NO, FACNO,
CLOSURE_DATE). To see every facility that could be active in the look-back window,
the range must reach back by the longest facility tenor as well:

    To   = reporting period end                                  (report_period, REPORT_MODE)
    From = period start - EXISTING_LOOKBACK_MONTHS - HISTORY_MAX_TENOR_MONTHS

    python branch_disb_data.py                       # mode from REPORT_MODE (orchestrator)
    python branch_disb_data.py --from-date 2019-01-01 --to-date 2026-11-15   # manual range

Saves to BASE_DIR/data/BRANCH_DISB_OLD_data/BRANCH DISBURSEMENT OLD_<from>_to_<to>_<stamp>.xlsx
"""
import sys
import argparse
import pyautogui
import time
import os
import shutil
import psutil

from datetime import datetime
from pathlib import Path
from dateutil.relativedelta import relativedelta
from dotenv import load_dotenv

from report_period import get_mode, get_period
from log_config import setup_logging

from igonore_firefox_updates import (
    dismiss_firefox_update_dialog,
    dismiss_save_password_prompt,
    oracle_sign_in2,
    search_bar_enter,
    set_date_field,
    click_button,
    click_toolbar_option,
    _wait_for_image,
    open_in_firefox,
)


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
DOWNLOADS_DIR = Path.home() / "Downloads"
REPORT_PREFIX = "BRANCH DISBURSEMENT OLD"
# Staging folder read by approved_facility_summary.py (moved to reports/<date>/data after a run)
REPORT_OUTPUT_DIR = BASE_DIR / "data" / "BRANCH_DISB_OLD_data"

MAX_ATTEMPTS = 3
DOWNLOAD_TIMEOUT = 1200          # multi-year range: allow 20 min

load_dotenv(BASE_DIR / "credentials.env")
# Existing-customer look-back (must match EXISTING_LOOKBACK_MONTHS in the summary)
EXISTING_LOOKBACK_MONTHS = int(os.environ.get("EXISTING_LOOKBACK_MONTHS", "3"))
# Longest facility tenor: a facility disbursed earlier than this has closed before the
# look-back window starts, so it cannot make anyone an Existing Customer.
HISTORY_MAX_TENOR_MONTHS = int(os.environ.get("HISTORY_MAX_TENOR_MONTHS", "84"))

log = setup_logging("branch_disb_data")


def history_range(from_date=None, to_date=None):
    """
    Return (first_day, last_day, tag) for the history download.

    Default (from REPORT_MODE / REPORT_N_MONTHS, same period as the Approval Report):
        last_day  = period end
        first_day = period start - EXISTING_LOOKBACK_MONTHS - HISTORY_MAX_TENOR_MONTHS
    --from-date / --to-date (YYYY-MM-DD) override either end.
    """
    mode = get_mode()
    period_start, period_end = get_period(mode)
    first = (datetime.strptime(from_date, "%Y-%m-%d").date() if from_date else
             period_start - relativedelta(months=EXISTING_LOOKBACK_MONTHS
                                          + HISTORY_MAX_TENOR_MONTHS))
    last = datetime.strptime(to_date, "%Y-%m-%d").date() if to_date else period_end
    if last < first:
        raise ValueError(f"To date {last} is before From date {first}")
    if not from_date and first > period_start - relativedelta(months=EXISTING_LOOKBACK_MONTHS):
        raise ValueError("History From date does not cover the look-back window")
    tag = f"{first:%Y-%m-%d}_to_{last:%Y-%m-%d}"
    return first, last, tag


def image_path(*parts):
    """Build an absolute path to an image under the script directory."""
    return str(BASE_DIR.joinpath(*parts))


# ============================================================
# DOWNLOAD UTILITIES
# ============================================================

def cleanup_orphaned_temp_files(prefix=REPORT_PREFIX):
    """Remove leftover partial-download files."""
    if not DOWNLOADS_DIR.exists():
        return

    for pattern in (f"{prefix}*.part", f"{prefix}*.crdownload"):
        for file_path in DOWNLOADS_DIR.glob(pattern):
            try:
                file_path.unlink()
                log.info(f"Removed: {file_path.name}")
            except OSError as exc:
                log.warning(f"Could not remove {file_path.name}: {exc}")


def wait_for_download(prefix, started_after, timeout=DOWNLOAD_TIMEOUT):
    """
    Wait for a new, nonempty download whose size remains stable.
    """
    if not DOWNLOADS_DIR.exists():
        raise FileNotFoundError(
            f"Downloads directory does not exist: {DOWNLOADS_DIR}"
        )

    deadline = time.time() + timeout
    previous_sizes = {}

    while time.time() < deadline:
        candidates = []

        for file_path in DOWNLOADS_DIR.glob(f"{prefix}*"):
            try:
                if not file_path.is_file():
                    continue

                if file_path.name.endswith((".part", ".crdownload")):
                    continue

                stat = file_path.stat()
                created_at = getattr(stat, "st_birthtime", stat.st_ctime)

                if created_at < started_after:
                    continue

                candidates.append(file_path)

            except OSError:
                continue

        candidates.sort(key=lambda path: path.stat().st_mtime, reverse=True)

        for file_path in candidates:
            try:
                part_file = file_path.with_name(file_path.name + ".part")
                crdownload_file = file_path.with_name(file_path.name + ".crdownload")

                if part_file.exists() or crdownload_file.exists():
                    log.info(f"Still downloading: {file_path.name}")
                    continue

                size = file_path.stat().st_size
                if size <= 0:
                    continue

                if previous_sizes.get(file_path) == size:
                    log.info(f"Download complete: {file_path.name} ({size:,} bytes)")
                    return file_path

                previous_sizes[file_path] = size
                log.info(f"Checking file size: {file_path.name} ({size:,} bytes)")

            except OSError:
                continue

        time.sleep(5)

    raise TimeoutError(
        f"Download did not complete within {timeout} seconds. "
        f"Check {DOWNLOADS_DIR} for incomplete files."
    )


def move_report_file(started_after, range_tag):
    """Move the completed report into the report output directory."""
    REPORT_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    log.info("Waiting for the report download...")

    downloaded_file = wait_for_download(
        prefix=REPORT_PREFIX,
        started_after=started_after,
        timeout=DOWNLOAD_TIMEOUT,
    )

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    destination = REPORT_OUTPUT_DIR / (
        f"{REPORT_PREFIX}_{range_tag}_{timestamp}{downloaded_file.suffix}"
    )

    counter = 1
    while destination.exists():
        destination = REPORT_OUTPUT_DIR / (
            f"{REPORT_PREFIX}_{range_tag}_{timestamp}_{counter}{downloaded_file.suffix}"
        )
        counter += 1

    try:
        shutil.move(str(downloaded_file), str(destination))
    except OSError as exc:
        raise RuntimeError(f"Could not move report to {destination}: {exc}") from exc

    if (
        downloaded_file.exists()
        or not destination.exists()
        or destination.stat().st_size == 0
    ):
        raise RuntimeError(
            "Report move verification failed: source still exists, "
            "or destination is missing or empty."
        )

    log.info(f"Report saved ({destination.stat().st_size:,} bytes): {destination}")
    return destination


# ============================================================
# GENERAL UTILITIES
# ============================================================

def _require(ok, step):
    """Raise an exception if a UI automation step fails."""
    if not ok:
        raise RuntimeError(f"UI step failed: {step}")
    log.info(f"UI step OK: {step}")


def close_browser():
    """Close Firefox processes gracefully, then terminate survivors."""
    processes = []

    for process in psutil.process_iter(["name"]):
        try:
            name = process.info.get("name")
            if name and name.lower() == "firefox.exe":
                processes.append(process)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

    if not processes:
        log.info("No Firefox process found to close.")
        return

    for process in processes:
        try:
            process.terminate()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

    _, alive = psutil.wait_procs(processes, timeout=10)

    for process in alive:
        try:
            process.kill()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

    log.info(f"Closed Firefox ({len(processes)} process(es)).")


def take_screenshot(attempt):
    """Save a screenshot after a failed attempt."""
    screenshot_path = BASE_DIR / f"branch_disb_old_fail_attempt_{attempt}.png"

    try:
        screenshot_path.parent.mkdir(parents=True, exist_ok=True)
        pyautogui.screenshot().save(str(screenshot_path))
        log.info(f"Saved: {screenshot_path}")
    except Exception as exc:
        log.warning(f"Could not save screenshot: {exc}")


# ============================================================
# REPORT SEARCH RESULT SELECTION
# ============================================================

def open_branch_disb_old_report(
    title_image,
    folder_image,
    open_image,
    timeout=15,
    confidence=0.7,
    search_pad=20,
):
    """
    Open BRANCH DISBURSEMENT OLD using the visible title and folder path.
    """

    # 1. Find the Title on screen
    title_loc = _wait_for_image(title_image, timeout, confidence)
    if not title_loc:
        log.error(f"Report title not found: {title_image}")
        return False

    log.info("Found report title. Looking for folder path underneath...")

    # 2. Define a region right below the title to find the folder path
    folder_search_region = (
        max(int(title_loc.left) - search_pad, 0),
        int(title_loc.top + title_loc.height - 5), # Start slightly above bottom edge
        600,                                       # Wide enough to capture the folder path
        50                                         # Height covering the folder line
    )

    # Pass region as a keyword argument
    folder_loc = _wait_for_image(
        folder_image, 
        timeout=5, 
        confidence=confidence, 
        region=folder_search_region
    )

    if not folder_loc:
        log.error(f"Folder path was not found under the title: {folder_image}")
        return False

    log.info("Matched title and folder path. Searching for Open button...")

    # 3. Define a region below the folder path to find the "Open" button
    open_search_region = (
        max(int(folder_loc.left) - search_pad, 0),
        int(folder_loc.top + folder_loc.height - 5),
        300,
        50
    )

    # 4. Locate and click the Open button
    start = time.time()
    while time.time() - start < timeout:
        try:
            open_loc = pyautogui.locateOnScreen(
                open_image, 
                confidence=confidence, 
                grayscale=True, 
                region=open_search_region
            )
            if open_loc:
                pyautogui.click(pyautogui.center(open_loc))
                log.info("Clicked Open for BRANCH DISBURSEMENT OLD.")
                return True
        except pyautogui.ImageNotFoundException:
            pass
        except Exception as exc:
            log.warning(f"Error locating Open link: {exc}")
            
        time.sleep(0.5)

    log.error("Open link was not found below the matched folder.")
    return False


# ============================================================
# DOWNLOAD RECOVERY
# ============================================================

def recover_download_if_onscreen(
    report_name_image,
    excel_icon_image,
    download_icon_image,
    download_dropdown_excel_image,
    timeout=10,
    confidence=0.7,
):
    """If the report renders in the browser rather than downloading, retry download."""
    banner_visible = _wait_for_image(report_name_image, timeout, confidence)

    if not banner_visible:
        log.info("Report banner not detected; continuing to wait for the download.")
        return False

    log.info("[RECOVER] Report appears on screen. Attempting to trigger its download again.")

    if _wait_for_image(excel_icon_image, timeout=3, confidence=confidence):
        if click_toolbar_option(
            excel_icon_image,
            download_dropdown_excel_image,
            timeout=timeout,
            confidence=confidence,
        ):
            log.info("[RECOVER] Excel toolbar option selected.")
            return True

    if click_toolbar_option(
        download_icon_image,
        download_dropdown_excel_image,
        timeout=timeout,
        confidence=confidence,
    ):
        log.info("[RECOVER] Download toolbar option selected.")
        return True

    log.warning("[RECOVER] Report banner found, but no download toolbar option could be selected.")
    return False


# ============================================================
# MAIN AUTOMATION
# ============================================================

def run_process(from_date=None, to_date=None):
    """Run the BRANCH DISBURSEMENT OLD download once for the history range."""

    os.chdir(BASE_DIR)

    credentials_file = BASE_DIR / "credentials.env"
    load_dotenv(dotenv_path=credentials_file)

    login_url = os.getenv("ORACLE_URL")
    username = os.getenv("ORACLE_USERNAME")
    password = os.getenv("ORACLE_PASSWORD")

    if not all((login_url, username, password)):
        raise RuntimeError(
            f"Missing ORACLE_URL, ORACLE_USERNAME, or ORACLE_PASSWORD in {credentials_file}"
        )

    search_text = REPORT_PREFIX

    # Image paths
    oracle_sign_in_page = image_path("img", "branch_disb", "Oracle_Signin_Page.png")
    oracle_signed_in_page = image_path("img", "branch_disb", "Half_UI.png")
    search_bar_image = image_path("img", "branch_disb", "Search_Bar.png")

    result_title_image = image_path("img", "branch_disb", "search_result", "Report_Title.png")
    result_folder_image = image_path("img", "branch_disb", "search_result", "Report_Folder_Path2.png")
    open_report_button_image = image_path("img", "branch_disb", "Open_Branch_Disb_Old_Report.png")

    from_date_field_image = image_path("img", "branch_disb", "From_Date.png")
    to_date_field_image = image_path("img", "branch_disb", "To_Date.png")
    apply_button_image = image_path("img", "branch_disb", "Apply.png")

    download_icon_image = image_path("img", "branch_disb", "Download_Icon.png")
    download_excel_option_image = image_path("img", "branch_disb", "Excel_File_Type.png")

    report_banner_image = image_path("img", "branch_disb", "previous", "report_name_img.png")
    excel_button_icon_image = image_path("img", "branch_disb", "previous", "download_button_img.png")

    save_password_prompt_image = image_path("img", "branch_disb", "Save_Password_Prompt.png")
    not_now_button_image = image_path("img", "branch_disb", "Not_Now_Button.png")

    # --------------------------------------------------------
    # DATE CALCULATION
    # --------------------------------------------------------
    first_day, last_day, range_tag = history_range(from_date, to_date)
    from_date = first_day.strftime("%m-%d-%Y")       # Oracle BI prompt format
    to_date = last_day.strftime("%m-%d-%Y")

    log.info(f"Report: {REPORT_PREFIX} | mode {get_mode()} | disbursed {first_day:%d-%b-%Y} "
             f"to {last_day:%d-%b-%Y} (look-back {EXISTING_LOOKBACK_MONTHS}m + "
             f"max tenor {HISTORY_MAX_TENOR_MONTHS}m)")
    log.info(f"From date: {from_date} | To date: {to_date}")

    # --------------------------------------------------------
    # CLEANUP
    # --------------------------------------------------------
    cleanup_orphaned_temp_files()

    # --------------------------------------------------------
    # OPEN ORACLE
    # --------------------------------------------------------
    # Firefox explicitly (never the default browser)
    open_in_firefox(login_url)

    time.sleep(7)

    update_image = image_path("img", "branch_disb", "Update_Message.png")
    dismiss_image = image_path("img", "branch_disb", "Dismiss_Button.png")

    dismiss_firefox_update_dialog(update_image, dismiss_image, timeout=15, confidence=0.8)

    # --------------------------------------------------------
    # LOGIN
    # --------------------------------------------------------
    _require(
        oracle_sign_in2(
            username,
            password,
            signin_page_image=oracle_sign_in_page,
            signed_in_image=oracle_signed_in_page,
            signed_in_timeout=5,
            timeout=15,
            confidence=0.7,
        ),
        "Oracle sign-in",
    )

    if Path(save_password_prompt_image).exists():
        dismiss_save_password_prompt(
            save_prompt_image=save_password_prompt_image,
            not_now_image=not_now_button_image,
            timeout=3,
            confidence=0.8,
        )
    else:
        pyautogui.press('escape')
        time.sleep(0.3)

    time.sleep(1)

    # --------------------------------------------------------
    # SEARCH
    # --------------------------------------------------------
    _require(
        search_bar_enter(search_text, search_bar_image, timeout=15, confidence=0.7),
        "Search for BRANCH DISBURSEMENT OLD",
    )

    time.sleep(1)

    # --------------------------------------------------------
    # SELECT REPORT
    # --------------------------------------------------------
    _require(
        open_branch_disb_old_report(
            title_image=result_title_image,
            folder_image=result_folder_image,
            open_image=open_report_button_image,
            timeout=15,
            confidence=0.7,
        ),
        "Open BRANCH DISBURSEMENT OLD report",
    )

    time.sleep(1)

    # --------------------------------------------------------
    # SET DATES
    # --------------------------------------------------------
    _require(
        set_date_field(from_date_field_image, from_date, timeout=5, confidence=0.7),
        "Set report From date",
    )

    time.sleep(1)

    _require(
        set_date_field(to_date_field_image, to_date, timeout=5, confidence=0.7),
        "Set report To date",
    )

    time.sleep(1)

    _require(
        click_button(apply_button_image, timeout=15, confidence=0.7),
        "Apply report date filters",
    )

    time.sleep(2)

    # --------------------------------------------------------
    # DOWNLOAD
    # --------------------------------------------------------
    cleanup_orphaned_temp_files()
    started_after = time.time()

    _require(
        click_toolbar_option(
            download_icon_image, 
            download_excel_option_image, 
            timeout=15, 
            confidence=0.7
        ),
        "Download report as Excel",
    )

    time.sleep(1)

    message_image = image_path("img", "branch_disb", "Message.png")
    button_image = image_path("img", "branch_disb", "Button.png")
    dismiss_firefox_update_dialog(message_image, button_image)

    recover_download_if_onscreen(
        report_name_image=report_banner_image,
        excel_icon_image=excel_button_icon_image,
        download_icon_image=download_icon_image,
        download_dropdown_excel_image=download_excel_option_image,
        timeout=10,
        confidence=0.7,
    )

    # --------------------------------------------------------
    # WAIT FOR DOWNLOAD AND MOVE REPORT
    # --------------------------------------------------------
    return move_report_file(started_after, range_tag)


# ============================================================
# ENTRY POINT AND RETRIES
# ============================================================

def parse_args():
    ap = argparse.ArgumentParser(
        description="Download BRANCH DISBURSEMENT OLD (customer history) from Oracle BI")
    ap.add_argument("--from-date", help="override From (YYYY-MM-DD); default: period start "
                                        "- look-back - max tenor")
    ap.add_argument("--to-date", help="override To (YYYY-MM-DD); default: period end")
    return ap.parse_args()


def main():
    args = parse_args()
    try:
        history_range(args.from_date, args.to_date)   # validates both and their order
    except ValueError as exc:
        log.error(f"Invalid history date range: {exc}")
        sys.exit(1)
    run_start = time.time()

    success = False
    destination = None

    for attempt in range(1, MAX_ATTEMPTS + 1):
        log.info(f"Attempt {attempt} of {MAX_ATTEMPTS}")

        try:
            destination = run_process(args.from_date, args.to_date)
            success = True
            log.info(f"Report saved at: {destination}")
            break
    
        except Exception as exc:
            log.error(f"Attempt {attempt} failed: {exc}", exc_info=True)
            take_screenshot(attempt)
            close_browser()
    
            if attempt < MAX_ATTEMPTS:
                log.info("Retrying in 5 seconds...")
                time.sleep(5)
            else:
                log.error("All attempts failed.")

    close_browser()

    if success:
        log.info(f"BRANCH DISBURSEMENT OLD completed successfully in "
                 f"{time.time() - run_start:.0f}s: {destination}")
        print(f"OUTPUT_FILE : {destination}")    # parsed by main_orchestrator.py
        sys.exit(0)

    log.error("BRANCH DISBURSEMENT OLD failed after all attempts.")
    sys.exit(1)


if __name__ == "__main__":
    main()