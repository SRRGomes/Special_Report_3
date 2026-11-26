"""
Orchestrator for the three reports. Every report runs the same 3 steps:

  1. Downloads   run each downloader script            [skipped by --skip-download]
  2. Summary     run the report's summary script -> PDF + Excel in reports/<YYYY-MM-DD>/
  3. Email       send the PDF + Excel                   [skipped by --no-emails]
  A failed step stops the run and (unless --no-emails) emails APPR_AF_ERROR_TO.

The reports differ only in their period, downloads, summary script and email group:

  --report approved-facility (default)   Approved Facility_(AF_3W_2W)_Analysis
      Period: --mode (daily / weekly / cumulative / monthly / past-n-months / custom)
      Downloads: Approval Report HO + Branch, Credit Score, RBP, BDO history
      Email: the group of the mode (DAILY_GROUP, CUMULATIVE_GROUP, ...)

  --report rental-shortfall              Rental Shortfall Report (runs on the last day of the month)
      Period: the latest 3 months of the LENDING csvs; as at the last day of the latest month
      Downloads: Credit Score, RBP, BDO history, one BDO per month
      Email: THREE_MONTHS_RENTAL_ANALYSIS_GROUP

  --report rental-dpd                    Rental Shortfall & 90 DPD Analysis (runs on the 1st)
      Period: data month = month before the run month (--as-at-month YYYY-MM);
              as at the 1st of the run month; activations = 12 months ending the month before
      Downloads: Credit Score, RBP, BDO history, BDO for the activation window
      Email: RENTAL_90DPD_ANALYSIS_GROUP
  LENDING csvs / parquet and the DPD file are shared manually (see data_request_alert.py).

    python main_orchestrator.py --mode cumulative
    python main_orchestrator.py --mode custom --from-date 2026-10-01 --to-date 2026-10-31
    python main_orchestrator.py --report rental-shortfall
    python main_orchestrator.py --report rental-dpd --as-at-month 2026-10
    python main_orchestrator.py --report rental-dpd --skip-download --no-emails   # testing

Exit code 0 on success, 1 on failure.
"""
import os
import sys
import argparse
import logging
import subprocess
import time
from datetime import timedelta
from pathlib import Path

from dotenv import load_dotenv

from report_period import (MODES, add_period_args, apply_period_args, get_period,
                           mode_text as describe_mode, period_label)
from log_config import log_file_path, setup_logging

log = logging.getLogger("main_orchestrator")

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / "credentials.env")

from approved_facility_summary import REPORT_NAME  # noqa: E402  (single place for the name)
from email_config import send_failure, send_report  # noqa: E402

REPORT_CHOICES = ("approved-facility", "rental-shortfall", "rental-dpd")

# Download steps: (name shown in the log, [script, args...])
STEP_APPROVAL_HO = ("Approval Report HO", ["approval_report_data.py", "--type", "HO"])
STEP_APPROVAL_BRANCH = ("Approval Report Branch", ["approval_report_data.py", "--type", "BRANCHES"])
STEP_CREDIT_SCORE = ("Credit Score Decision Report",
                     ["erp_decision_report_data.py", "--report", "credit-score"])
STEP_RBP = ("RBP Security Approved", ["erp_decision_report_data.py", "--report", "rbp"])
STEP_BDO_HISTORY = ("Branch Disbursement Old (customer history)", ["branch_disb_data.py"])


def step_bdo_range(label, first_day, last_day):
    """BRANCH DISBURSEMENT OLD for a fixed date range."""
    return (f"Branch Disbursement Old ({label})",
            ["branch_disb_data.py", "--from-date", first_day.isoformat(),
             "--to-date", last_day.isoformat()])


# Kept for scripts that import them
DOWNLOAD_STEPS = [STEP_APPROVAL_HO, STEP_APPROVAL_BRANCH, STEP_CREDIT_SCORE, STEP_RBP,
                  STEP_BDO_HISTORY]
SUMMARY_SCRIPT = "approved_facility_summary.py"
RENTAL_SUMMARY_SCRIPT = "rental_shortfall_summary.py"
DPD_SUMMARY_SCRIPT = "rental_dpd_summary.py"

STEP_TIMES = []          # (step, status, seconds) for the end-of-run summary


# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------

def run_step(name, args, env):
    """
    Run a child script, echo its output, return (exit code, output lines).
    Child scripts write to the same daily log file themselves, so their output is
    printed here rather than logged again.
    """
    log.info(f"----- STEP: {name} | {' '.join(args)} -----")
    t0 = time.time()
    proc = subprocess.Popen([sys.executable, str(BASE_DIR / args[0]), *args[1:]],
                            cwd=BASE_DIR, env=env, stdout=subprocess.PIPE,
                            stderr=subprocess.STDOUT, text=True, encoding="utf-8",
                            errors="replace")
    lines = []
    for line in proc.stdout:
        print(line, end="", flush=True)
        lines.append(line.rstrip("\n"))
    proc.wait()
    secs = time.time() - t0
    status = "OK" if proc.returncode == 0 else f"FAILED (exit {proc.returncode})"
    STEP_TIMES.append((name, status, secs))
    (log.info if proc.returncode == 0 else log.error)(f"STEP {name}: {status} in {secs:.0f}s")
    return proc.returncode, lines


def parse_outputs(lines):
    """RECORDS / AMOUNT / PERIOD / PDF / EXCEL lines printed by a summary script."""
    out = {}
    for line in lines:
        for key in ("RECORDS", "AMOUNT", "PERIOD", "PDF", "EXCEL"):
            if line.startswith(key) and ":" in line:
                out[key] = line.split(":", 1)[1].strip()
    return out


def log_run_summary(run_start, ok, report_name=REPORT_NAME):
    for name, status, secs in STEP_TIMES:
        log.info(f"  {name:<32} {status:<20} {secs:>6.0f}s")
    (log.info if ok else log.error)(
        f"===== {report_name} run {'COMPLETED' if ok else 'FAILED'} in "
        f"{time.time() - run_start:.0f}s | log: {log_file_path()} =====")


def run_pipeline(args, name, report, mode, mode_text, env, steps, summary_args):
    """
    The 3 steps shared by every report. Exits the program (0 = OK, 1 = failed).
      report     'approved-facility' / 'rental-shortfall' / 'rental-dpd' (email names, group)
      mode       report mode passed to send_report (selects the group for approved-facility)
    """
    run_start = time.time()

    def fail(step):
        if not args.no_emails:
            send_failure(step, mode_text, report=report)
        log_run_summary(run_start, ok=False, report_name=name)
        sys.exit(1)

    # 1. Downloads
    if args.skip_download:
        log.info("--skip-download: using the latest downloaded files.")
    else:
        for step_name, step in steps:
            code, _ = run_step(step_name, step, env)
            if code != 0:
                log.error(f"{step_name} download failed; stopping. See {log_file_path().name} "
                          f"and the screenshots in the project folder.")
                fail(step_name)

    # 2. Summary -> PDF + Excel
    if args.history_file:
        summary_args = summary_args + ["--history-file", args.history_file]
    code, lines = run_step(f"Summary ({name})", summary_args, env)
    outputs = parse_outputs(lines)
    if code != 0 or "PDF" not in outputs or "EXCEL" not in outputs:
        log.error("Summary step failed (no PDF/EXCEL in its output); no email sent.")
        fail("Summary")
    log.info(f"Summary: {outputs.get('RECORDS')} AF/3W/2W facilities, "
             f"LKR {float(outputs.get('AMOUNT', 0)):,.0f} | {outputs.get('PERIOD')}")

    # 3. Email
    if args.no_emails:
        log.info("--no-emails: email skipped.")
    else:
        try:
            send_report(outputs["PDF"], outputs["EXCEL"], outputs.get("PERIOD", ""), mode,
                        mode_text, outputs.get("RECORDS"), outputs.get("AMOUNT"), report=report)
        except Exception as e:
            log.error(f"Email failed: {e}", exc_info=True)
            log_run_summary(run_start, ok=False, report_name=name)
            sys.exit(1)

    log.info(f"PDF  : {outputs['PDF']}")
    log.info(f"Excel: {outputs['EXCEL']}")
    log_run_summary(run_start, ok=True, report_name=name)
    sys.exit(0)


# ------------------------------------------------------------
# THE THREE REPORTS (period, steps, summary script)
# ------------------------------------------------------------

def run_approved_facility(args):
    mode_text = describe_mode(args.mode, args.n_months)
    start, end = get_period(args.mode, args.n_months)
    log.info(f"{REPORT_NAME} | mode {mode_text} | {period_label(start, end)} "
             f"({start} to {end}) | skip-download={args.skip_download} "
             f"| no-emails={args.no_emails}")

    env = dict(os.environ, REPORT_MODE=args.mode, REPORT_N_MONTHS=str(args.n_months),
               PYTHONIOENCODING="utf-8")
    summary_args = [SUMMARY_SCRIPT, "--mode", args.mode, "--n-months", str(args.n_months)]
    if args.mode == "custom":
        summary_args += ["--from-date", start.isoformat(), "--to-date", end.isoformat()]
    run_pipeline(args, REPORT_NAME, "approved-facility", args.mode, mode_text, env,
                 DOWNLOAD_STEPS, summary_args)


def run_rental(args):
    """Rental Shortfall Report: months from the LENDING csvs."""
    import rental_shortfall_summary as rss
    name = rss.RENTAL_REPORT_TITLE
    lending_dir = Path(args.lending_dir) if args.lending_dir else rss.LENDING_DIR
    try:
        start, end, months = rss.lending_period(lending_dir)
    except Exception as e:
        log.error(f"{name}: cannot read the LENDING files in {lending_dir}: {e}", exc_info=True)
        log_run_summary(time.time(), ok=False, report_name=name)
        sys.exit(1)
    mode_text = f"{len(months)} months ({', '.join(m.strftime('%b-%Y') for m in months)})"
    log.info(f"{name} | {mode_text} | {start} to {end} | skip-download={args.skip_download} "
             f"| no-emails={args.no_emails}")

    # Child scripts see the LENDING months as a custom period; the history download
    # uses the 12-month look-back of the v2 customer rule.
    env = dict(os.environ, REPORT_MODE="custom", REPORT_N_MONTHS=str(len(months)),
               REPORT_FROM_DATE=start.isoformat(), REPORT_TO_DATE=end.isoformat(),
               EXISTING_LOOKBACK_MONTHS=str(rss.RENTAL_LOOKBACK_MONTHS),
               PYTHONIOENCODING="utf-8")
    steps = [STEP_CREDIT_SCORE, STEP_RBP, STEP_BDO_HISTORY]
    steps += [step_bdo_range(f"{m.strftime('%b-%Y')} activations",
                             m.start_time.date(), m.end_time.date()) for m in months]
    summary_args = [RENTAL_SUMMARY_SCRIPT, "--lending-dir", str(lending_dir)]
    run_pipeline(args, name, "rental-shortfall", "rental-shortfall", mode_text, env,
                 steps, summary_args)


def run_rental_dpd(args):
    """Rental Shortfall & 90 DPD Analysis: data month = month before the run month."""
    import rental_dpd_summary as rds
    name = rds.DPD_REPORT_TITLE
    lending_dir = Path(args.lending_dir) if args.lending_dir else rds.LENDING_DIR
    try:
        start, act_end, as_at, as_at_month, window = rds.dpd_period(args.as_at_month)
    except Exception as e:
        log.error(f"{name}: invalid --as-at-month {args.as_at_month}: {e}", exc_info=True)
        log_run_summary(time.time(), ok=False, report_name=name)
        sys.exit(1)
    mode_text = (f"activations {window[0].strftime('%b-%Y')} to {window[1].strftime('%b-%Y')}, "
                 f"as at {as_at_month.strftime('%b-%Y')}")
    log.info(f"{name} | {mode_text} | skip-download={args.skip_download} "
             f"| no-emails={args.no_emails}")

    # Child scripts see first activation day .. run date (1st after the data month) as a
    # custom period: decisions are pulled from 3 months (look-back) before it, and the
    # history reaches the run date for the Existing test "open on the run date".
    run_date = as_at + timedelta(days=1)
    env = dict(os.environ, REPORT_MODE="custom", REPORT_N_MONTHS=str(rds.ACTIVATION_MONTHS),
               REPORT_FROM_DATE=start.isoformat(), REPORT_TO_DATE=run_date.isoformat(),
               EXISTING_LOOKBACK_MONTHS=str(rds.rss.RENTAL_LOOKBACK_MONTHS),
               PYTHONIOENCODING="utf-8")
    steps = [STEP_CREDIT_SCORE, STEP_RBP, STEP_BDO_HISTORY,
             step_bdo_range(f"activations {start:%b-%Y} to {act_end:%b-%Y}", start, act_end)]
    summary_args = [DPD_SUMMARY_SCRIPT, "--lending-dir", str(lending_dir),
                    "--as-at-month", str(as_at_month)]
    if args.dpd_file:
        summary_args += ["--dpd-file", args.dpd_file]
    run_pipeline(args, name, "rental-dpd", "rental-dpd", mode_text, env, steps, summary_args)


# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------

def parse_args():
    ap = argparse.ArgumentParser(description=f"{REPORT_NAME} / rental reports orchestrator")
    ap.add_argument("--report", choices=REPORT_CHOICES, default="approved-facility",
                    help="which report to run (default: approved-facility)")
    ap.add_argument("--mode", choices=MODES, default="cumulative",
                    help="approved-facility: reporting period")
    ap.add_argument("--n-months", type=int, default=3, help="months for past-n-months")
    add_period_args(ap)                       # --from-date / --to-date for --mode custom
    ap.add_argument("--lending-dir", help="rental reports: LENDING folder "
                                          "(default: data/LENDING_data)")
    ap.add_argument("--dpd-file", help="rental-dpd: DPD file (default: newest in data/DPD_data)")
    ap.add_argument("--as-at-month", help="rental-dpd: YYYY-MM (default: month before the run month)")
    ap.add_argument("--lookback-months", type=int,
                    help="months of credit decisions before the period start "
                         "(default 3, or DECISION_LOOKBACK_MONTHS in credentials.env)")
    ap.add_argument("--history-file", help="BRANCH DISBURSEMENT OLD file for the Existing/New "
                                           "check (default: latest downloaded)")
    ap.add_argument("--skip-download", action="store_true",
                    help="testing: use the latest files already downloaded")
    ap.add_argument("--no-emails", action="store_true", help="testing: do not send emails")
    args = ap.parse_args()

    if args.lookback_months is not None:
        if args.lookback_months < 0:
            ap.error("--lookback-months must be 0 or more")
        # Copied into every child's environment -> erp_decision_report_data.py
        os.environ["DECISION_LOOKBACK_MONTHS"] = str(args.lookback_months)
    if args.report != "approved-facility":
        if args.from_date or args.to_date or args.mode != "cumulative":
            ap.error(f"--report {args.report} sets its own period; "
                     "do not pass --mode / --from-date / --to-date")
    else:
        apply_period_args(ap, args)           # validates; sets REPORT_FROM/TO_DATE for children
    return args


def main():
    args = parse_args()
    setup_logging("main_orchestrator")
    if args.report == "rental-shortfall":
        run_rental(args)
    elif args.report == "rental-dpd":
        run_rental_dpd(args)
    else:
        run_approved_facility(args)


if __name__ == "__main__":
    main()
