"""
Shared browser helpers for the Oracle BI downloaders (approval_report_data.py,
branch_disb_data.py) and the ERP downloader.

  find_firefox / open_in_firefox     always open Firefox, never the default browser
  dismiss_firefox_update_dialog      close the "update Firefox" pop-up if it appears
  dismiss_save_password_prompt       close the "save password" pop-up
  oracle_sign_in / oracle_sign_in2   Oracle BI login (skipped if already signed in)
  search_bar_enter, open_catalog_report, click_field, set_date_field, click_button,
  click_toolbar_option, handle_report_download
                                     find a screen element by its image and click / type
  _wait_for_image                    the image search used by all of them
                                     (confidence 0.7, grayscale)
"""
import os
import time
import shutil
import logging
import subprocess
import pyautogui

# Messages go to the caller's logging setup (console + daily log file)
log = logging.getLogger("firefox_ui")


# ----------------------------------------------------------------------
# Firefox location (do NOT rely on webbrowser / the default browser:
# on machines where Firefox is not registered with Python, webbrowser
# falls back to the default browser, usually Edge)
# ----------------------------------------------------------------------
def find_firefox():
    """
    Return the full path to firefox.exe, or None.
    Order: FIREFOX_PATH env var (credentials.env) -> Windows registry (App Paths)
    -> standard install folders -> PATH.
    """
    env_path = os.environ.get("FIREFOX_PATH", "").strip().strip('"')
    if env_path:
        if os.path.isfile(env_path):
            return env_path
        log.warning(f"FIREFOX_PATH is set but not found: {env_path}")

    try:
        import winreg
        key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\firefox.exe"
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(hive, key_path) as key:
                    path = winreg.QueryValue(key, None).strip('"')
                    if os.path.isfile(path):
                        return path
            except OSError:
                continue
    except ImportError:
        pass                                    # not Windows

    candidates = [
        os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"),
                     "Mozilla Firefox", "firefox.exe"),
        os.path.join(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
                     "Mozilla Firefox", "firefox.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Mozilla Firefox", "firefox.exe"),
    ]
    for path in candidates:
        if path and os.path.isfile(path):
            return path

    return shutil.which("firefox") or shutil.which("firefox.exe")


def open_in_firefox(url):
    """Open url in Firefox (never the default browser). Raises if Firefox is not found."""
    firefox = find_firefox()
    if not firefox:
        raise FileNotFoundError(
            "Firefox not found. Install Firefox or set FIREFOX_PATH in credentials.env, "
            r"e.g. FIREFOX_PATH=C:\Program Files\Mozilla Firefox\firefox.exe")
    subprocess.Popen([firefox, url])
    log.info(f"Opened {url} in Firefox ({firefox})")
    return firefox

# Capture the update dialog (when it appears)
# pyautogui.screenshot('firefox_update_dialog.png', region=(x, y, width, height))

# Capture the dismiss button
# pyautogui.screenshot('dismiss_button.png', region=(x, y, width, height))

# Or use this helper:
def capture_templates():
    print("Position mouse over the UPDATE DIALOG area and press Enter")
    input()
    x, y = pyautogui.position()
    pyautogui.screenshot('firefox_update_dialog.png', region=(x-200, y-100, 400, 200))
    
    print("Position mouse over the DISMISS BUTTON and press Enter")
    input()
    x, y = pyautogui.position()
    pyautogui.screenshot('dismiss_button.png', region=(x-40, y-15, 80, 30))
    print("Templates captured!")

def dismiss_firefox_update_dialog(update_dialog_image='firefox_update_dialog.png', 
                                   dismiss_button_image='dismiss_button.png', 
                                   timeout=10, 
                                   confidence=0.7):
    """
    First check for Firefox update dialog, then find and click Dismiss button if found.
    
    Args:
        update_dialog_image: Path to saved image of the Firefox update dialog
        dismiss_button_image: Path to saved image of the 'Dismiss' button
        timeout: Maximum seconds to wait for dialog to appear
        confidence: Matching confidence (0-1)
    
    Returns:
        bool: True if dialog was found and dismissed, False if dialog not found
    """
    start_time = time.time()
    
    while time.time() - start_time < timeout:
        try:
            # Step 1: Check for update dialog box
            dialog_location = pyautogui.locateOnScreen(
                update_dialog_image, 
                confidence=confidence,
                grayscale=True
            )
            
            if dialog_location:
                log.info("Firefox update dialog detected")
                
                # Step 2: Look for Dismiss button
                try:
                    button_location = pyautogui.locateOnScreen(
                        dismiss_button_image,
                        confidence=confidence,
                        grayscale=True
                    )
                    
                    if button_location:
                        # Step 3: Click Dismiss button
                        center = pyautogui.center(button_location)
                        pyautogui.click(center)
                        log.info("Dismiss button clicked successfully")
                        return True
                    else:
                        log.warning("Dialog found but Dismiss button not located")
                        return False
                        
                except pyautogui.ImageNotFoundException:
                    log.warning("Dismiss button not found in dialog")
                    return False
                    
        except pyautogui.ImageNotFoundException:
            # Dialog not found yet, continue scanning
            pass
        except Exception as e:
            log.warning(f"Update dialog check error: {e}")
        
        # Wait before next scan
        time.sleep(0.5)
    
    log.info(f"No Firefox update dialog detected within {timeout}s (nothing to dismiss)")
    return False

# ============================================================
# FIREFOX SAVE-PASSWORD PROMPT
# ============================================================

def dismiss_save_password_prompt(save_prompt_image='save_password_prompt.png',
                                 not_now_image='not_now_button.png',
                                 timeout=3,
                                 confidence=0.8):
    """
    Dismiss Firefox's 'Save password?' doorhanger if it appears.
    """
    if not _wait_for_image(save_prompt_image, timeout, confidence):
        log.info("No save-password prompt detected.")
        return False

    log.info("Firefox save-password prompt detected; dismissing.")

    if click_button(not_now_image, timeout=5, confidence=confidence):
        time.sleep(0.3)
        return True

    log.warning("'Not now' button not found; pressing Escape")
    pyautogui.press('escape')
    time.sleep(0.2)
    return True

def oracle_sign_in(username, password,
                   signin_page_image='Oracle_SignIn.png',
                   signed_in_image='Oracle_SignedIn.png',
                   signed_in_timeout=5,
                   timeout=15, confidence=0.7,
                   userid_frac=(0.50, 0.28),
                   password_frac=(0.50, 0.40),
                   signin_frac=(0.50, 0.60),
                   use_tab=True):
    """
    If already signed in, skip the login workflow entirely. Otherwise recognize
    the sign-in page and log in.

    signed_in_image: a static crop unique to the signed-in UI (something that
        never appears on the login screen — e.g. a toolbar, the catalog header,
        the user/account menu).
    signed_in_timeout: short wait for the signed-in UI before falling back to
        login. Keep it small so a genuine login page isn't delayed much.
    """
    # Short-circuit: are we already in?
    if signed_in_image:
        if _wait_for_image(signed_in_image, signed_in_timeout, confidence):
            log.info("Already signed in — skipping login")
            return True

    # Not signed in: proceed with login
    loc = _wait_for_image(signin_page_image, timeout, confidence)
    if not loc:
        log.warning("Sign-in page not detected")
        return False

    def _point(frac):
        return (int(loc.left + frac[0] * loc.width),
                int(loc.top  + frac[1] * loc.height))

    # User ID
    pyautogui.click(_point(userid_frac))
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.typewrite(username, interval=0.03)

    # Password
    if use_tab:
        pyautogui.press('tab')
    else:
        pyautogui.click(_point(password_frac))
        pyautogui.hotkey('ctrl', 'a')
    pyautogui.typewrite(password, interval=0.03)

    # Sign In
    pyautogui.click(_point(signin_frac))
    log.info("Submitted sign-in")
    return True

# ============================================================
# ORACLE SIGN-IN
# ============================================================
def oracle_sign_in2(username, password,
                   signin_page_image='Oracle_SignIn.png',
                   signed_in_image='Oracle_SignedIn.png',
                   signed_in_timeout=5,
                   timeout=15, confidence=0.7,
                   userid_frac=(0.50, 0.28),
                   password_frac=(0.50, 0.40),
                   signin_frac=(0.50, 0.60),
                   use_tab=True):
    """
    If already signed in, skip the login workflow entirely. Otherwise recognize
    the sign-in page and log in.

    signed_in_image: a static crop unique to the signed-in UI (something that
        never appears on the login screen — e.g. a toolbar, the catalog header,
        the user/account menu).
    signed_in_timeout: short wait for the signed-in UI before falling back to
        login. Keep it small so a genuine login page isn't delayed much.
    """
    # Short-circuit: are we already in?
    if signed_in_image:
        if _wait_for_image(signed_in_image, signed_in_timeout, confidence):
            log.info("Already signed in — skipping login")
            return True

    # Not signed in: proceed with login
    loc = _wait_for_image(signin_page_image, timeout, confidence)
    if not loc:
        log.warning("Sign-in page not detected")
        return False

    def _point(frac):
        return (int(loc.left + frac[0] * loc.width),
                int(loc.top  + frac[1] * loc.height))

    # User ID
    pyautogui.click(_point(userid_frac))
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.typewrite(username, interval=0.03)

    # Password
    if use_tab:
        pyautogui.press('tab')
    else:
        pyautogui.click(_point(password_frac))
        pyautogui.hotkey('ctrl', 'a')
    pyautogui.typewrite(password, interval=0.03)

    # Sign In
    pyautogui.click(_point(signin_frac))
    log.info("Submitted sign-in")
    return True

def search_bar_enter(text, search_bar_image='Search_Bar.png',
                     timeout=10, confidence=0.7,
                     x_offset=0, y_offset=0,
                     clear_first=True, submit=True):
    """
    Locate the search bar by its image, click into it, and type `text`.

    x_offset/y_offset: nudge off-center — the crop includes the magnifier icon
    on the right, so a small negative x_offset lands in the input, not the icon.
    submit: press Enter after typing.
    """
    loc = _wait_for_image(search_bar_image, timeout, confidence)
    if not loc:
        log.warning(f"Search bar not found: {search_bar_image}")
        return False
    cx, cy = pyautogui.center(loc)
    pyautogui.click(cx + x_offset, cy + y_offset)
    if clear_first:
        pyautogui.hotkey('ctrl', 'a')
    pyautogui.typewrite(text, interval=0.03)
    if submit:
        pyautogui.press('enter')
    log.info(f"Entered search text: {text}")
    return True

def open_catalog_report(report_image='Credit_Committee_Decision_Report.png',
                        open_image='Open_Credit_Committee_Decision_Report.png',
                        timeout=15, confidence=0.7,
                        search_pad=10):
    """
    Recognize a catalog report entry, then click *its* Open link.

    Because every row on the page has an 'Open' link, we first locate the report
    block, then search for the Open image only within a region bounded by that
    block (its own area plus a little padding below), so we click the right one.
    """
    report_loc = _wait_for_image(report_image, timeout, confidence)
    if not report_loc:
        log.warning(f"Report entry not found: {report_image}")
        return False
    log.info(f"Found report entry: {report_image}")

    # Restrict the Open search to the report block's own footprint (+pad below,
    # where the Open/More links sit). region = (left, top, width, height).
    region = (max(int(report_loc.left) - search_pad, 0),
              max(int(report_loc.top) - search_pad, 0),
              int(report_loc.width) + 2 * search_pad,
              int(report_loc.height) + search_pad * 4)

    start = time.time()
    while time.time() - start < timeout:
        try:
            open_loc = pyautogui.locateOnScreen(
                open_image, confidence=confidence, grayscale=True, region=region)
            if open_loc:
                pyautogui.click(pyautogui.center(open_loc))
                log.info("Clicked Open on the report")
                return True
        except pyautogui.ImageNotFoundException:
            pass
        except Exception as e:
            log.warning(f"Error locating Open: {e}")
        time.sleep(0.5)

    log.warning("Open link not found within the report region")
    return False

def _wait_for_image(image, timeout=10, confidence=0.7, grayscale=True, region=None):
    """Poll the screen (or `region` = (left, top, width, height)) for `image`. Returns the box or None."""
    start = time.time()
    while time.time() - start < timeout:
        try:
            loc = pyautogui.locateOnScreen(image, confidence=confidence, grayscale=grayscale,
                                           region=region)
            if loc:
                return loc
        except pyautogui.ImageNotFoundException:
            pass
        except Exception as e:
            log.warning(f"Error locating {image}: {e}")
        time.sleep(0.5)
    return None


# 1. Text fields (From Date / To Date)
def click_field(field_image, timeout=10, confidence=0.7, x_offset=0, y_offset=0, clicks=1):
    """
    Locate a text field by its image and click it. Use x_offset/y_offset to land
    inside the input box rather than on the label. Returns True if clicked.
    """
    loc = _wait_for_image(field_image, timeout, confidence)
    if not loc:
        log.warning(f"Field not found: {field_image}")
        return False
    cx, cy = pyautogui.center(loc)
    pyautogui.click(cx + x_offset, cy + y_offset, clicks=clicks)
    log.info(f"Clicked field: {field_image}")
    return True


def set_date_field(field_image, date_text, timeout=10, confidence=0.7,
                   x_offset=0, y_offset=0, clear_first=True):
    """Click a date field and type into it. clear_first does Ctrl+A before typing."""
    if not click_field(field_image, timeout, confidence, x_offset, y_offset):
        return False
    if clear_first:
        pyautogui.hotkey('ctrl', 'a')
    pyautogui.typewrite(date_text, interval=0.03)
    log.info(f"Set {field_image} = {date_text}")
    return True



# 2. Button (Apply, Open, etc.)
def click_button(button_image, timeout=10, confidence=0.7):
    """Locate a button by its image and click its center. Returns True if clicked."""
    loc = _wait_for_image(button_image, timeout, confidence)
    if not loc:
        log.warning(f"Button not found: {button_image}")
        return False
    pyautogui.click(pyautogui.center(loc))
    log.info(f"Clicked button: {button_image}")
    return True


# 3. Toolbar icon + dropdown option (Download -> Excel *.xlsx)
def click_toolbar_option(icon_image, option_image, timeout=10, confidence=0.7, menu_delay=0.5):
    """
    Click a toolbar icon to open its dropdown, then click the given option.
    Returns True if the option was clicked.
    """
    icon_loc = _wait_for_image(icon_image, timeout, confidence)
    if not icon_loc:
        log.warning(f"Toolbar icon not found: {icon_image}")
        return False
    pyautogui.click(pyautogui.center(icon_loc))
    log.info(f"Clicked toolbar icon: {icon_image}")

    time.sleep(menu_delay)  # let dropdown render

    option_loc = _wait_for_image(option_image, timeout, confidence)
    if not option_loc:
        log.warning(f"Dropdown option not found: {option_image}")
        return False
    pyautogui.click(pyautogui.center(option_loc))
    log.info(f"Clicked option: {option_image}")
    return True

# 4. Report download: click Excel option -> wait for "leave page?" -> click Leave
def handle_report_download(excel_option_image, message_image, leave_button_image,
                           report_timeout=180, confidence=0.7,
                           option_timeout=15, grayscale=True):
    """
    Finish a BI report export after the download dropdown is already open.

    Steps:
      1. Click the "Excel (*.xlsx)" option in the open dropdown.
      2. Wait up to report_timeout for the browser's "leave page?" dialog
         (it only appears once the export has been generated, which can be slow).
      3. Click the "Leave page" button.

    Does NOT click the generate/download toolbar icon — do that before calling
    this (the dropdown must already be open).

    Returns True if Leave was clicked, False if any step's image is not found.
    """
    # 1. Excel option
    opt = _wait_for_image(excel_option_image, option_timeout, confidence, grayscale=grayscale)
    if not opt:
        log.warning(f"Excel option not found: {excel_option_image}")
        return False
    pyautogui.click(pyautogui.center(opt))
    log.info(f"Clicked Excel option: {excel_option_image}")

    # 2. Wait for the "leave page?" dialog (allow full report_timeout to render)
    msg = _wait_for_image(message_image, report_timeout, confidence, grayscale=grayscale)
    if not msg:
        log.warning(f"'Leave page' dialog not found within {report_timeout}s: {message_image}")
        return False
    log.info("'Leave page' dialog detected")

    # 3. Leave page button
    btn = _wait_for_image(leave_button_image, option_timeout, confidence, grayscale=grayscale)
    if not btn:
        log.warning(f"'Leave page' button not found: {leave_button_image}")
        return False
    pyautogui.click(pyautogui.center(btn))
    log.info("Clicked 'Leave page' — download starting")
    return True

# Usage example:
# dismiss_firefox_update_dialog()