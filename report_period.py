"""
Shared reporting-period logic for the Approved Facility (AF / 3W / 2W) pipeline.

The mode is read from the REPORT_MODE environment variable (set by
main_orchestrator.py for every child script); past-n-months reads N from
REPORT_N_MONTHS; custom reads REPORT_FROM_DATE / REPORT_TO_DATE (YYYY-MM-DD).

  daily         : yesterday
  weekly        : last 7 days, ending yesterday
  cumulative    : 1st of this month to yesterday (on the 1st: the whole previous month)
  monthly       : previous full month
  past-n-months : N full months, ending last month
  custom        : --from-date to --to-date (both required, both inclusive)

Each downloader converts the period to its own format:
Oracle BI MM-DD-YYYY, SmartOps ERP YYYY-MM-DD.
"""
import os
from datetime import date, datetime, timedelta
from dateutil.relativedelta import relativedelta

MODES = ("daily", "weekly", "cumulative", "monthly", "past-n-months", "custom")
# YYYY-MM-DD only: 01-10-2026 would be ambiguous (1 Oct vs 10 Jan)
CLI_DATE_FORMATS = ("%Y-%m-%d",)
DEFAULT_MODE = "cumulative"
DEFAULT_N_MONTHS = 3


def get_mode():
    mode = os.environ.get("REPORT_MODE", DEFAULT_MODE).strip().lower()
    if mode not in MODES:
        raise ValueError(f"Unknown REPORT_MODE '{mode}'. Use one of: {', '.join(MODES)}")
    return mode


def get_n_months():
    raw = os.environ.get("REPORT_N_MONTHS", str(DEFAULT_N_MONTHS)).strip()
    n = int(raw)
    if n < 1:
        raise ValueError(f"REPORT_N_MONTHS must be >= 1 (got {raw})")
    return n


def parse_cli_date(text):
    """'2026-10-01' -> date. Raises ValueError otherwise."""
    text = (text or "").strip()
    for fmt in CLI_DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"Invalid date '{text}': use YYYY-MM-DD (e.g. 2026-10-01)")


def get_custom_dates(today=None):
    """(from, to) for custom mode from REPORT_FROM_DATE / REPORT_TO_DATE."""
    raw_from = os.environ.get("REPORT_FROM_DATE", "").strip()
    raw_to = os.environ.get("REPORT_TO_DATE", "").strip()
    missing = [n for n, v in (("--from-date", raw_from), ("--to-date", raw_to)) if not v]
    if missing:
        raise ValueError(f"custom mode needs {' and '.join(missing)}")
    start, end = parse_cli_date(raw_from), parse_cli_date(raw_to)
    if start > end:
        raise ValueError(f"--from-date {start} is after --to-date {end}")
    if end > (today or date.today()):
        raise ValueError(f"--to-date {end} is in the future")
    return start, end


def add_period_args(ap):
    """Add --from-date / --to-date (custom mode) to an ArgumentParser."""
    ap.add_argument("--from-date", help="custom mode: start date, YYYY-MM-DD")
    ap.add_argument("--to-date", help="custom mode: end date, YYYY-MM-DD")


def apply_period_args(ap, args):
    """
    Validate --mode / --from-date / --to-date and publish custom dates to the
    environment (REPORT_FROM_DATE / REPORT_TO_DATE, ISO) so every child script
    sees the same period. Exits with a usage error on bad input.
    """
    given = args.from_date or args.to_date
    if args.mode != "custom":
        if given:
            ap.error("--from-date / --to-date are only used with --mode custom")
        return
    if not (args.from_date and args.to_date):
        ap.error("--mode custom needs both --from-date and --to-date")
    try:
        start, end = parse_cli_date(args.from_date), parse_cli_date(args.to_date)
    except ValueError as e:
        ap.error(str(e))
    os.environ["REPORT_FROM_DATE"] = start.isoformat()
    os.environ["REPORT_TO_DATE"] = end.isoformat()
    try:
        get_custom_dates()
    except ValueError as e:
        ap.error(str(e))


def mode_text(mode, n_months=None):
    """Human label: 'cumulative', '3 months', 'custom range'."""
    if mode == "past-n-months":
        return f"{n_months or get_n_months()} months"
    if mode == "custom":
        return "custom range"
    return mode


def mode_tag(mode, n_months=None):
    """File-name part: 'cumulative', 'past-3-months', 'custom'."""
    if mode == "past-n-months":
        return f"past-{n_months or get_n_months()}-months"
    return mode


def get_period(mode=None, n_months=None, today=None):
    """(start, end) dates, both inclusive."""
    mode = mode or get_mode()
    today = today or date.today()
    yesterday = today - timedelta(days=1)

    if mode == "custom":
        return get_custom_dates(today)

    if mode == "daily":
        return yesterday, yesterday
    if mode == "weekly":
        return yesterday - timedelta(days=6), yesterday
    if mode == "cumulative":
        return yesterday.replace(day=1), yesterday

    first_this_month = today.replace(day=1)
    n = 1 if mode == "monthly" else (n_months or get_n_months())
    return first_this_month - relativedelta(months=n), first_this_month - timedelta(days=1)


def _is_month_end(d):
    return (d + timedelta(days=1)).day == 1


def period_label(start, end):
    """'12-Nov-2026' | 'October-2026' | 'August-2026 to October-2026' | '01-Nov-2026 to 12-Nov-2026'."""
    if start == end:
        return start.strftime("%d-%b-%Y")
    if start.day == 1 and _is_month_end(end):
        if (start.year, start.month) == (end.year, end.month):
            return start.strftime("%B-%Y")
        return f"{start:%B-%Y} to {end:%B-%Y}"
    return f"{start:%d-%b-%Y} to {end:%d-%b-%Y}"


def period_tag(start, end):
    """File-name-safe version of period_label."""
    return period_label(start, end).replace(" ", "_")
