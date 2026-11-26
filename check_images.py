"""
Check which pyautogui images are found on THIS PC's screen, and at what confidence.

Images captured on another PC often fail because of a different screen resolution,
Windows display scaling (Settings > System > Display > Scale) or Firefox zoom.

Usage:
  1. Open the Oracle BI page by hand so the screen shows what the images show
     (e.g. the BRANCH DISBURSEMENT OLD report with its From/To date fields).
  2. python check_images.py                      # all images under img/branch_disb
     python check_images.py img/branch_disb/From_Date.png img/branch_disb/Apply.png
     python check_images.py --folder img          # every image under img/
  3. Re-capture any image marked NOT FOUND (or found only below 0.7) on this PC
     with the Snipping Tool, same file name, cropped tightly around the element.

It does not click anything. You have 5 seconds to switch to the Firefox window.
"""
import sys
import time
import argparse
from pathlib import Path

import pyautogui
from PIL import Image

BASE_DIR = Path(__file__).resolve().parent
LEVELS = [0.95, 0.9, 0.85, 0.8, 0.75, 0.7, 0.65, 0.6, 0.55, 0.5]
SCRIPT_CONFIDENCE = 0.7          # what the downloaders use


def best_confidence(image):
    """Highest confidence level at which the image is found (grayscale), or None."""
    for level in LEVELS:
        try:
            box = pyautogui.locateOnScreen(str(image), confidence=level, grayscale=True)
        except pyautogui.ImageNotFoundException:
            box = None
        if box:
            return level, box
    return None, None


def main():
    ap = argparse.ArgumentParser(description="Check pyautogui images against the current screen")
    ap.add_argument("images", nargs="*", help="image files (default: --folder)")
    ap.add_argument("--folder", default="img/branch_disb", help="folder searched recursively")
    ap.add_argument("--delay", type=int, default=5, help="seconds before checking")
    args = ap.parse_args()

    images = [Path(p) for p in args.images] or sorted((BASE_DIR / args.folder).rglob("*.png"))
    if not images:
        sys.exit(f"No .png files found in {BASE_DIR / args.folder}")

    w, h = pyautogui.size()
    print(f"Screen: {w}x{h} (pyautogui). Checking {len(images)} image(s) in {args.delay}s - "
          f"switch to the Firefox window now.")
    time.sleep(args.delay)

    failed = 0
    for img in images:
        try:
            iw, ih = Image.open(img).size
        except Exception:
            iw = ih = "?"
        level, box = best_confidence(img)
        if level is None:
            status = "NOT FOUND (even at 0.5)"
            failed += 1
        elif level < SCRIPT_CONFIDENCE:
            status = f"found only at {level} (scripts use {SCRIPT_CONFIDENCE}) -> re-capture"
            failed += 1
        else:
            status = f"OK at {level}  at x={box.left}, y={box.top}"
        try:
            name = img.relative_to(BASE_DIR)
        except ValueError:
            name = img
        print(f"  {str(name):<60} {iw}x{ih}  {status}")

    print(f"\n{len(images) - failed} OK, {failed} to re-capture. Note: an image is only found if "
          f"that element is visible on the screen right now.")


if __name__ == "__main__":
    main()
