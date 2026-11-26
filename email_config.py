"""
Email for the Approved Facility (AF / 3W / 2W) Analysis.

Used by main_orchestrator.py, or on its own to (re)send the latest report
from reports/<YYYY-MM-DD>/:

    python email_config.py --mode cumulative
    python email_config.py --mode past-n-months --n-months 3 --date 2026-11-13
    python email_config.py --mode daily --dry-run          # show what would be sent
    python email_config.py --mode custom --from-date 2026-10-01 --to-date 2026-10-31
    python email_config.py --report rental-shortfall --dry-run   # latest Rental Shortfall Report
    python email_config.py --report rental-dpd --dry-run         # latest Rental Shortfall & 90 DPD Analysis

credentials.env:
    SMTP_SERVER, SMTP_PORT (25), SMTP_USE_TLS (false), SMTP_USERNAME, SMTP_PASSWORD, EMAIL_FROM
    Recipient group per mode (comma/semicolon separated):
        daily          -> DAILY_GROUP
        cumulative     -> CUMULATIVE_GROUP
        monthly        -> MONTHLY_GROUP
        past-n-months  -> N_MONTHS_GROUP
        custom         -> CUSTOM_GROUP
        weekly         -> WEEKLY_GROUP   (no group defined yet; set it to email weekly runs)
    Rental Shortfall Report (--report rental-shortfall) -> THREE_MONTHS_RENTAL_ANALYSIS_GROUP
    Rental Shortfall & 90 DPD Analysis (--report rental-dpd) -> RENTAL_90DPD_ANALYSIS_GROUP
    APPR_AF_ERROR_TO                    optional, failure alerts
    DATA_PROVIDER                       data-request alerts (data_request_alert.py)
"""
import os
import sys
import glob
import argparse
import logging
import smtplib
from pathlib import Path
from datetime import datetime
from email.message import EmailMessage

from dotenv import load_dotenv

from report_period import (MODES, add_period_args, apply_period_args, get_period,
                           mode_tag, mode_text as describe_mode, period_label, period_tag)
from log_config import setup_logging

# Module logger; handlers are set up by the caller (orchestrator) or in main()
log = logging.getLogger("email_config")

BASE_DIR = Path(__file__).resolve().parent
REPORTS_DIR = BASE_DIR / "reports"
load_dotenv(BASE_DIR / "credentials.env")

from approved_facility_summary import REPORT_NAME  # noqa: E402  (single place for the name)

# Report mode -> recipient group in credentials.env
MODE_GROUPS = {
    "daily": "DAILY_GROUP",
    "weekly": "WEEKLY_GROUP",
    "cumulative": "CUMULATIVE_GROUP",
    "monthly": "MONTHLY_GROUP",
    "past-n-months": "N_MONTHS_GROUP",
    "custom": "CUSTOM_GROUP",
}
ERROR_TO_VAR = "APPR_AF_ERROR_TO"

# Rental Shortfall Report: one group, whatever the months
RENTAL_GROUP = "THREE_MONTHS_RENTAL_ANALYSIS_GROUP"
# Rental Shortfall & 90 DPD Analysis: one group
DPD_GROUP = "RENTAL_90DPD_ANALYSIS_GROUP"


def _rental():
    import rental_shortfall_summary as rss       # only when that report is emailed
    return rss


def _dpd():
    import rental_dpd_summary as rds            # only when that report is emailed
    return rds


def report_names(report):
    """(email subject name, recipient group or None = by mode) for a report."""
    if report == "rental-dpd":
        return _dpd().DPD_EMAIL_SUBJECT, DPD_GROUP
    if report == "rental-shortfall":
        return _rental().RENTAL_EMAIL_SUBJECT, RENTAL_GROUP
    return REPORT_NAME, None

# ============================================================
# RECIPIENTS / SMTP
# ============================================================

def _split(raw):
    return [a.strip() for a in (raw or "").replace(";", ",").split(",") if a.strip()]


def recipients(var):
    return _split(os.environ.get(var))


def group_var(mode):
    """credentials.env variable holding the recipients for this mode."""
    if mode not in MODE_GROUPS:
        raise ValueError(f"No recipient group defined for mode '{mode}'")
    return MODE_GROUPS[mode]


def send_email(subject, html, to, cc=None, attachments=()):
    server = os.environ.get("SMTP_SERVER")
    sender = os.environ.get("EMAIL_FROM")
    if not server or not sender:
        raise EnvironmentError("SMTP_SERVER and EMAIL_FROM must be set in credentials.env")
    if not to:
        raise EnvironmentError("No recipients given (check the recipient group in credentials.env)")

    msg = EmailMessage()
    msg["Subject"], msg["From"], msg["To"] = subject, sender, ", ".join(to)
    if cc:
        msg["Cc"] = ", ".join(cc)
    msg.set_content("This email requires an HTML-capable mail client.")
    msg.add_alternative(html, subtype="html")
    for path in attachments:
        path = Path(path)
        if not path.exists():
            raise FileNotFoundError(f"Attachment not found: {path}")
        subtype = ("pdf" if path.suffix.lower() == ".pdf"
                   else "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        msg.add_attachment(path.read_bytes(), maintype="application", subtype=subtype,
                           filename=path.name)
        log.info(f"Attached {path.name} ({path.stat().st_size:,} bytes)")

    port = int(os.environ.get("SMTP_PORT", "25"))
    use_tls = os.environ.get("SMTP_USE_TLS", "false").lower() in ("1", "true", "yes")
    log.info(f"Connecting to SMTP {server}:{port} (TLS {'on' if use_tls else 'off'}) as {sender}")
    with smtplib.SMTP(server, port, timeout=60) as smtp:
        if use_tls:
            smtp.starttls()
        user, pwd = os.environ.get("SMTP_USERNAME"), os.environ.get("SMTP_PASSWORD")
        if user and pwd:
            smtp.login(user, pwd)
            log.info(f"SMTP login OK ({user})")
        smtp.send_message(msg)          # To + Cc taken from the headers
    log.info(f"Sent '{subject}' | To: {', '.join(to)} | Cc: {', '.join(cc or []) or '-'}")

# ============================================================
# MESSAGES
# ============================================================

def report_message(period, mode_text, records=None, amount=None, report="approved-facility"):
    name, _ = report_names(report)
    subject = f"{name} - {period}"
    what = ("Activated facilities" if report in ("rental-shortfall", "rental-dpd")
            else "Approved facilities")
    stats = ""
    if records is not None and amount is not None:
        stats = (f"<p>{what} (AF, 3W &amp; 2W): <b>{int(float(records)):,}</b> "
                 f"| Lease Amount: <b>LKR {float(amount):,.0f}</b></p>")
    html = f"""
    <p>Dear All,</p>
    <p>Please find attached the <b>{name}</b> for <b>{period}</b> ({mode_text}).</p>
    {stats}
    <p>This is an automated email.</p>"""
    return subject, html


def send_report(pdf, excel, period, mode, mode_text, records=None, amount=None, dry_run=False,
                report="approved-facility"):
    subject, html = report_message(period, mode_text, records, amount, report)
    var = report_names(report)[1] or group_var(mode)
    to = recipients(var)
    log.info(f"Recipient group for mode '{mode}': {var} ({len(to)} address(es))")
    if dry_run:
        log.info(f"[DRY-RUN] Subject: {subject}")
        log.info(f"[DRY-RUN] To ({var}): {to or '(none - set ' + var + ' in credentials.env)'}")
        log.info(f"[DRY-RUN] Attachments: {Path(pdf).name}, {Path(excel).name}")
        return
    if not to:
        raise EnvironmentError(f"{var} is empty or missing in credentials.env "
                               f"(recipients for {mode} runs)")
    send_email(subject, html, to, attachments=[pdf, excel])


def send_failure(step, mode_text, report="approved-facility"):
    """Failure alert to APPR_AF_ERROR_TO (skipped if not set). Never raises."""
    to = recipients(ERROR_TO_VAR)
    if not to:
        log.info(f"{ERROR_TO_VAR} not set; no failure email sent")
        return
    log.info(f"Sending failure alert for step '{step}'")
    name, _ = report_names(report)
    subject = f"[FAILED] {name} - {step}"
    html = (f"<p>The {name} run ({mode_text}) failed at <b>{step}</b>. "
            f"Check the orchestrator log on the automation machine.</p>")
    try:
        send_email(subject, html, to)
    except Exception as e:
        log.warning(f"Could not send failure email: {e}")

# ============================================================
# STANDALONE: send the latest report in reports/<date>/
# ============================================================

def latest_outputs(mode, n_months, run_date):
    """Newest PDF + Excel for this mode/period in reports/<run_date>/."""
    start, end = get_period(mode, n_months)
    stem = f"{REPORT_NAME}_{mode_tag(mode, n_months)}_{period_tag(start, end)}"
    folder = REPORTS_DIR / run_date
    found = {}
    for ext in (".pdf", ".xlsx"):
        files = [Path(f) for f in glob.glob(glob.escape(str(folder / stem)) + ext)]
        if not files:
            raise FileNotFoundError(f"{stem}{ext} not found in {folder}")
        found[ext] = max(files, key=os.path.getmtime)
    return found[".pdf"], found[".xlsx"], period_label(start, end)


def latest_rental_outputs(run_date, report="rental-shortfall"):
    """Newest Rental Shortfall Report (or Rental Shortfall & 90 DPD Analysis) PDF + Excel."""
    name = _dpd().DPD_FILE_NAME if report == "rental-dpd" else _rental().RENTAL_FILE_NAME
    folder = REPORTS_DIR / run_date
    found = {}
    for ext in (".pdf", ".xlsx"):
        files = [Path(f) for f in glob.glob(glob.escape(str(folder / name)) + "_*" + ext)]
        if not files:
            raise FileNotFoundError(f"{name}_*{ext} not found in {folder}")
        found[ext] = max(files, key=os.path.getmtime)
    period = found[".pdf"].stem[len(name) + 1:].replace("_", " ")
    return found[".pdf"], found[".xlsx"], period


def main():
    ap = argparse.ArgumentParser(description=f"Email the {REPORT_NAME} or Rental Shortfall Report")
    ap.add_argument("--report", choices=("approved-facility", "rental-shortfall", "rental-dpd"),
                    default="approved-facility")
    ap.add_argument("--mode", choices=MODES, default="cumulative")
    ap.add_argument("--n-months", type=int, default=3)
    add_period_args(ap)                       # --from-date / --to-date for --mode custom
    ap.add_argument("--date", default=datetime.now().strftime("%Y-%m-%d"),
                    help="run date folder reports/<YYYY-MM-DD> (default: today)")
    ap.add_argument("--dry-run", action="store_true", help="print, do not send")
    args = ap.parse_args()
    if args.report in ("rental-shortfall", "rental-dpd"):
        setup_logging("email_config")
        try:
            pdf, excel, period = latest_rental_outputs(args.date, args.report)
            log.info(f"Latest outputs in reports/{args.date}: {pdf.name}, {excel.name}")
            send_report(pdf, excel, period, args.report,
                        "3 months" if args.report == "rental-shortfall" else "12 months activation",
                        dry_run=args.dry_run, report=args.report)
        except Exception as e:
            log.error(f"Email failed: {e}", exc_info=True)
            sys.exit(1)
        return
    apply_period_args(ap, args)
    setup_logging("email_config")
    try:
        pdf, excel, period = latest_outputs(args.mode, args.n_months, args.date)
        log.info(f"Latest outputs in reports/{args.date}: {pdf.name}, {excel.name}")
        mode_text = describe_mode(args.mode, args.n_months)
        send_report(pdf, excel, period, args.mode, mode_text, dry_run=args.dry_run)
    except Exception as e:
        log.error(f"Email failed: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
