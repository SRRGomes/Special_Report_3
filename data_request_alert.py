"""
Data-request alert: 2 days before a rental report runs, email DATA_PROVIDER
(credentials.env) the manual files that report needs, the date filter for each,
where to put them, and whether they are already there.

Run days (fixed):
  Rental Shortfall Report              last day of each month   -> alert 2 days before
  Rental Shortfall & 90 DPD Analysis   1st of each month        -> alert 2 days before
                                                                   (2nd-last day of the month)
The Approved Facility report needs no manual data, so it has no alert.

How to call it
  From scheduler.py, once a day:      python data_request_alert.py
      -> sends only the alerts that are due today, otherwise logs "nothing due".
  By hand (send now, for a given run date):
      python data_request_alert.py --report rental-shortfall --run-date 2026-11-30
      python data_request_alert.py --report rental-dpd --run-date 2026-12-01 --dry-run

  In Python (e.g. inside scheduler.py):
      import data_request_alert
      data_request_alert.send_due_alerts()          # today's alerts, if any
"""
import os
import sys
import argparse
import logging
from pathlib import Path
from datetime import date, datetime, timedelta

import pandas as pd
from dotenv import load_dotenv

from log_config import setup_logging

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / "credentials.env")

log = logging.getLogger("data_request_alert")

DATA_PROVIDER_VAR = "DATA_PROVIDER"     # recipients in credentials.env
ALERT_DAYS_BEFORE = 2
REPORTS = ("rental-shortfall", "rental-dpd")


# ------------------------------------------------------------
# 1. WHEN: run dates and alert dates
# ------------------------------------------------------------

def run_date_for(report, today):
    """The next run date of a report on or after `today`."""
    if report == "rental-shortfall":                       # last day of this month
        return (pd.Timestamp(today) + pd.offsets.MonthEnd(0)).date()
    first_next = (pd.Timestamp(today) + pd.offsets.MonthBegin(1)).date()
    return today if today.day == 1 else first_next         # 1st of a month


def due_reports(today=None):
    """Reports whose alert day (run date - 2 days) is today."""
    today = today or date.today()
    due = []
    for report in REPORTS:
        run = run_date_for(report, today + timedelta(days=ALERT_DAYS_BEFORE))
        if run - timedelta(days=ALERT_DAYS_BEFORE) == today:
            due.append((report, run))
    return due


# ------------------------------------------------------------
# 2. WHAT: the manual files each report needs
# ------------------------------------------------------------

def _month_text(m):
    return m.strftime("%B %Y")


def needs_rental_shortfall(run_date):
    """LENDING csv for the run month and the 2 months before (as at the run month end)."""
    import rental_shortfall_summary as rss
    import rental_dpd_summary as rds
    run_month = pd.Period(run_date, "M")
    months = list(pd.period_range(run_month - 2, run_month, freq="M"))
    on_hand = rds.lending_months_on_hand(rss.LENDING_DIR, include_parquet=False)
    rows = []
    for m in months:
        rows.append({
            "data": f"LENDING portfolio csv - {_month_text(m)}",
            "filter": f"MONTHEND = {m.end_time:%d-%m-%Y}",
            "folder": str(rss.LENDING_DIR),
            "status": f"Received ({on_hand[m]})" if m in on_hand else "Missing",
        })
    title = (f"as at {run_month.end_time:%d %B %Y}; months "
             f"{_month_text(months[0])} to {_month_text(months[-1])}")
    return title, rows


def needs_rental_dpd(run_date):
    """12 LENDING month-ends up to the as-at month + the DPD file of the as-at month."""
    import rental_dpd_summary as rds
    as_at_month = pd.Period(run_date, "M") - 1              # run on 1 Nov -> October data
    months = rds.needed_lending_months(as_at_month)
    on_hand = rds.lending_months_on_hand(rds.LENDING_DIR, include_parquet=True)
    rows = []
    for m in months:
        rows.append({
            "data": f"LENDING portfolio - {_month_text(m)}",
            "filter": f"MONTHEND = {m.end_time:%d-%m-%Y}",
            "folder": f"{rds.LENDING_DIR} (csv) or {rds.LENDING_DIR / rds.LENDING_BASE_SUBDIR} (parquet)",
            "status": f"Received ({on_hand[m]})" if m in on_hand else "Missing",
        })
    dpd_file = rds.dpd_file_with_month(as_at_month)
    rows.append({
        "data": "DPD file (FacNo, clientcode, Month, DPD_Adj)",
        "filter": f"Month = {_month_text(as_at_month)}",
        "folder": str(rds.DPD_DIR),
        "status": f"Received ({dpd_file})" if dpd_file else "Missing",
    })
    first, last = rds.activation_window(as_at_month)
    title = (f"as at {(as_at_month + 1).start_time:%d %B %Y}; activations "
             f"{_month_text(first)} to {_month_text(last)}")
    return title, rows


# ------------------------------------------------------------
# 3. SEND
# ------------------------------------------------------------

def build_email(report, run_date):
    if report == "rental-shortfall":
        import rental_shortfall_summary as rss
        name = rss.RENTAL_REPORT_TITLE
        title, rows = needs_rental_shortfall(run_date)
    else:
        import rental_dpd_summary as rds
        name = rds.DPD_REPORT_TITLE
        title, rows = needs_rental_dpd(run_date)

    missing = sum(r["status"] == "Missing" for r in rows)
    subject = (f"[Data request] {name} - runs {run_date:%d-%b-%Y}"
               + (f" ({missing} file(s) missing)" if missing else " (all received)"))
    cell = "border:1px solid #8EAADB;padding:4px 8px;font-size:12px"
    table = "".join(
        f"<tr><td style='{cell}'>{r['data']}</td><td style='{cell}'>{r['filter']}</td>"
        f"<td style='{cell}'>{r['folder']}</td>"
        f"<td style='{cell};color:{'#A32D2D' if r['status'] == 'Missing' else '#3B6D11'}'>"
        f"<b>{r['status']}</b></td></tr>" for r in rows)
    html = f"""
    <p>Dear Team,</p>
    <p>The <b>{name}</b> runs on <b>{run_date:%A, %d %B %Y}</b> ({title}).
    Please provide the following data with the date filter shown:</p>
    <table style="border-collapse:collapse">
      <tr style="background:#4472C4;color:#fff">
        <th style='{cell}'>Data</th><th style='{cell}'>Date filter</th>
        <th style='{cell}'>Put it in</th><th style='{cell}'>Status</th></tr>
      {table}
    </table>
    <p>This is an automated email.</p>"""
    return subject, html, rows


def send_alert(report, run_date, dry_run=False):
    from email_config import recipients, send_email
    subject, html, rows = build_email(report, run_date)
    for r in rows:
        log.info(f"  {r['status']:<10} {r['data']} | {r['filter']}")
    to = recipients(DATA_PROVIDER_VAR)
    if dry_run:
        log.info(f"[DRY-RUN] Subject: {subject} | To ({DATA_PROVIDER_VAR}): {to or '(not set)'}")
        return
    if not to:
        raise EnvironmentError(f"{DATA_PROVIDER_VAR} is empty or missing in credentials.env")
    send_email(subject, html, to)


def send_due_alerts(today=None, dry_run=False):
    """Send every alert due today. Returns the list of (report, run date) sent."""
    today = today or date.today()
    due = due_reports(today)
    if not due:
        log.info(f"No data-request alert due on {today:%d-%b-%Y}")
    for report, run in due:
        log.info(f"Alert due: {report} runs {run:%d-%b-%Y}")
        send_alert(report, run, dry_run)
    return due


def main():
    ap = argparse.ArgumentParser(description="Data-request alerts for the rental reports")
    ap.add_argument("--report", choices=REPORTS, help="send this report's alert now")
    ap.add_argument("--run-date", help="YYYY-MM-DD run date for --report "
                                       "(default: its next run date)")
    ap.add_argument("--today", help="YYYY-MM-DD: pretend today is this date (testing)")
    ap.add_argument("--dry-run", action="store_true", help="log the email, do not send")
    args = ap.parse_args()
    setup_logging("data_request_alert")
    try:
        today = datetime.strptime(args.today, "%Y-%m-%d").date() if args.today else date.today()
        if args.report:
            run = (datetime.strptime(args.run_date, "%Y-%m-%d").date() if args.run_date
                   else run_date_for(args.report, today))
            send_alert(args.report, run, args.dry_run)
        else:
            send_due_alerts(today, args.dry_run)
    except Exception as e:
        log.error(f"Data-request alert failed: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
