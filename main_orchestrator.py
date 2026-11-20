"""
Orchestrator — Approved Facility (AF / 3W / 2W) Analysis.

    python main_orchestrator.py --mode cumulative
    python main_orchestrator.py --mode past-n-months --n-months 3
    python main_orchestrator.py --mode custom --from-date 2026-10-01 --to-date 2026-10-31
    python main_orchestrator.py --mode daily --skip-download --no-emails      # testing
    python main_orchestrator.py --report rental-shortfall                       # Rental Shortfall Report

Steps:
  1. Download Approval Report HO + Branch and BRANCH DISBURSEMENT OLD (Oracle BI) and the
     Credit Score Decision / RBP Security Approved reports (SmartOps ERP)
                                                             [skipped by --skip-download]
  2. approved_facility_summary.py -> PDF + Excel in reports/<YYYY-MM-DD>/
  3. Email the PDF + Excel via email_config.py              [skipped by --no-emails]

--report rental-shortfall (Rental Shortfall Report): the months come from the LENDING
portfolio csvs (data/LENDING_data, MONTHEND), not from --mode. Downloads: Credit Score /
RBP (in <=90-day pieces), BRANCH DISBURSEMENT OLD history, and one BRANCH DISBURSEMENT
OLD per month (that month's activations). Email: THREE_MONTHS_RENTAL_ANALYSIS_GROUP.

Exit code 0 on success, 1 on failure.
"""
import os
import sys
import argparse
import logging
import subprocess
import time
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

DOWNLOAD_STEPS = [
    ("Approval Report HO", ["approval_report_data.py", "--type", "HO"]),
    ("Approval Report Branch", ["approval_report_data.py", "--type", "BRANCHES"]),
    ("Credit Score Decision Report", ["erp_decision_report_data.py", "--report", "credit-score"]),
    ("RBP Security Approved", ["erp_decision_report_data.py", "--report", "rbp"]),
    ("Branch Disbursement Old (customer history)", ["branch_disb_data.py"]),
]
SUMMARY_SCRIPT = "approved_facility_summary.py"

REPORT_CHOICES = ("approved-facility", "rental-shortfall")
RENTAL_SUMMARY_SCRIPT = "rental_shortfall_summary.py"
RENTAL_DECISION_STEPS = [s for s in DOWNLOAD_STEPS if s[1][0] == "erp_decision_report_data.py"]
RENTAL_HISTORY_STEP = ("Branch Disbursement Old (customer history)", ["branch_disb_data.py"])


STEP_TIMES = []          # (step, status, seconds) for the end-of-run summary


def run_step(name, args, env):
    """
    Run a child script, echo its output to the console, return (exit code, output lines).
    Child scripts write their own lines to the same daily log file, so their output is
    echoed (print) rather than logged again.
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
    out = {}
    for line in lines:
        for key in ("RECORDS", "AMOUNT", "PERIOD", "PDF", "EXCEL"):
            if line.startswith(key) and ":" in line:
                out[key] = line.split(":", 1)[1].strip()
    return out


# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description=f"{REPORT_NAME} / Rental Shortfall Report orchestrator")
    ap.add_argument("--report", choices=REPORT_CHOICES, default="approved-facility",
                    help="which report to run (default: approved-facility)")
    ap.add_argument("--lending-dir", help="rental-shortfall: LENDING csv folder "
                                          "(default: data/LENDING_data)")
    ap.add_argument("--mode", choices=MODES, default="cumulative")
    ap.add_argument("--n-months", type=int, default=3, help="months for past-n-months")
    add_period_args(ap)                       # --from-date / --to-date for --mode custom
    ap.add_argument("--skip-download", action="store_true",
                    help="testing: use the latest files already downloaded")
    ap.add_argument("--no-emails", action="store_true", help="testing: do not send emails")
    ap.add_argument("--history-file", help="BRANCH DISBURSEMENT OLD file for the Existing/New "
                                           "check (default: latest downloaded)")
    args = ap.parse_args()
    if args.report == "rental-shortfall":
        if args.from_date or args.to_date or args.mode != "cumulative":
            ap.error("--report rental-shortfall takes its months from the LENDING files; "
                     "do not pass --mode / --from-date / --to-date")
        setup_logging("main_orchestrator")
        run_rental(args)                      # exits
    apply_period_args(ap, args)               # validates; sets REPORT_FROM/TO_DATE for children
    setup_logging("main_orchestrator")
    run_start = time.time()

    mode_text = describe_mode(args.mode, args.n_months)
    start, end = get_period(args.mode, args.n_months)
    log.info(f"{REPORT_NAME} | mode {mode_text} | {period_label(start, end)} "
             f"({start} to {end}) | skip-download={args.skip_download} "
             f"| no-emails={args.no_emails}")

    env = dict(os.environ, REPORT_MODE=args.mode, REPORT_N_MONTHS=str(args.n_months),
               PYTHONIOENCODING="utf-8")

    if not args.skip_download:
        for name, step in DOWNLOAD_STEPS:
            code, _ = run_step(name, step, env)
            if code != 0:
                log.error(f"{name} download failed; stopping. "
                          f"See {log_file_path().name} and the report_fail_*.png screenshots.")
                if not args.no_emails:
                    send_failure(name, mode_text)
                log_run_summary(run_start, ok=False)
                sys.exit(1)
    else:
        log.info("--skip-download: using the latest downloaded files.")

    summary_args = [SUMMARY_SCRIPT, "--mode", args.mode, "--n-months", str(args.n_months)]
    if args.mode == "custom":
        summary_args += ["--from-date", start.isoformat(), "--to-date", end.isoformat()]
    if args.history_file:
        summary_args += ["--history-file", args.history_file]
    code, lines = run_step("Summary (PDF + Excel)", summary_args, env)
    outputs = parse_outputs(lines)
    if code != 0 or "PDF" not in outputs or "EXCEL" not in outputs:
        log.error("Summary step failed (no PDF/EXCEL in its output); no email sent.")
        if not args.no_emails:
            send_failure("Summary", mode_text)
        log_run_summary(run_start, ok=False)
        sys.exit(1)
    log.info(f"Summary: {outputs.get('RECORDS')} AF/3W/2W facilities, "
             f"LKR {float(outputs.get('AMOUNT', 0)):,.0f} | period {outputs.get('PERIOD')}")

    if args.no_emails:
        log.info("--no-emails: email skipped.")
    else:
        log.info("Sending report email")
        try:
            send_report(outputs["PDF"], outputs["EXCEL"], outputs.get("PERIOD", ""),
                        args.mode, mode_text, outputs.get("RECORDS"), outputs.get("AMOUNT"))
        except Exception as e:
            log.error(f"Email failed: {e}", exc_info=True)
            log_run_summary(run_start, ok=False)
            sys.exit(1)

    log.info(f"PDF  : {outputs['PDF']}")
    log.info(f"Excel: {outputs['EXCEL']}")
    log_run_summary(run_start, ok=True)
    sys.exit(0)


def log_run_summary(run_start, ok, report_name=REPORT_NAME):
    for name, status, secs in STEP_TIMES:
        log.info(f"  {name:<32} {status:<20} {secs:>6.0f}s")
    (log.info if ok else log.error)(
        f"===== {report_name} run {'COMPLETED' if ok else 'FAILED'} in "
        f"{time.time() - run_start:.0f}s | log: {log_file_path()} =====")


# ------------------------------------------------------------
# RENTAL SHORTFALL REPORT
# ------------------------------------------------------------

def run_rental(args):
    """--report rental-shortfall: months from LENDING, downloads, summary, email. Exits."""
    import rental_shortfall_summary as rss
    name = rss.RENTAL_REPORT_TITLE
    run_start = time.time()
    lending_dir = Path(args.lending_dir) if args.lending_dir else rss.LENDING_DIR
    try:
        start, end, months = rss.lending_period(lending_dir)
    except Exception as e:
        log.error(f"{name}: cannot read the LENDING files in {lending_dir}: {e}", exc_info=True)
        log_run_summary(run_start, ok=False, report_name=name)
        sys.exit(1)
    mode_text = f"{len(months)} months ({', '.join(m.strftime('%b-%Y') for m in months)})"
    log.info(f"{name} | {mode_text} | {start} to {end} | skip-download={args.skip_download} "
             f"| no-emails={args.no_emails}")

    # Every child script sees the LENDING months as a custom period;
    # the history download uses the 12-month look-back of the v2 customer rule.
    env = dict(os.environ, REPORT_MODE="custom", REPORT_N_MONTHS=str(len(months)),
               REPORT_FROM_DATE=start.isoformat(), REPORT_TO_DATE=end.isoformat(),
               EXISTING_LOOKBACK_MONTHS=str(rss.RENTAL_LOOKBACK_MONTHS),
               PYTHONIOENCODING="utf-8")

    steps = list(RENTAL_DECISION_STEPS) + [RENTAL_HISTORY_STEP]
    for m in months:
        steps.append((f"Branch Disbursement Old ({m.strftime('%b-%Y')} activations)",
                      ["branch_disb_data.py", "--from-date", m.start_time.strftime("%Y-%m-%d"),
                       "--to-date", m.end_time.strftime("%Y-%m-%d")]))

    def fail(step):
        if not args.no_emails:
            send_failure(step, mode_text, report="rental-shortfall")
        log_run_summary(run_start, ok=False, report_name=name)
        sys.exit(1)

    if not args.skip_download:
        for step_name, step in steps:
            code, _ = run_step(step_name, step, env)
            if code != 0:
                log.error(f"{step_name} download failed; stopping. See {log_file_path().name}.")
                fail(step_name)
    else:
        log.info("--skip-download: using the latest downloaded files.")

    summary_args = [RENTAL_SUMMARY_SCRIPT, "--lending-dir", str(lending_dir)]
    if args.history_file:
        summary_args += ["--history-file", args.history_file]
    code, lines = run_step(f"Summary ({name})", summary_args, env)
    outputs = parse_outputs(lines)
    if code != 0 or "PDF" not in outputs or "EXCEL" not in outputs:
        log.error("Summary step failed (no PDF/EXCEL in its output); no email sent.")
        fail("Summary")
    log.info(f"Summary: {outputs.get('RECORDS')} AF/3W/2W facilities, "
             f"LKR {float(outputs.get('AMOUNT', 0)):,.0f} | {outputs.get('PERIOD')}")

    if args.no_emails:
        log.info("--no-emails: email skipped.")
    else:
        try:
            send_report(outputs["PDF"], outputs["EXCEL"], outputs.get("PERIOD", ""),
                        "rental-shortfall", mode_text, outputs.get("RECORDS"),
                        outputs.get("AMOUNT"), report="rental-shortfall")
        except Exception as e:
            log.error(f"Email failed: {e}", exc_info=True)
            log_run_summary(run_start, ok=False, report_name=name)
            sys.exit(1)

    log.info(f"PDF  : {outputs['PDF']}")
    log.info(f"Excel: {outputs['EXCEL']}")
    log_run_summary(run_start, ok=True, report_name=name)
    sys.exit(0)


if __name__ == "__main__":
    main()
