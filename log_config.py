"""
Shared logging for the Approved Facility (AF / 3W / 2W) pipeline.

Every script logs to the console and to one daily file:
    logs/approved_facility_<YYYY-MM-DD>.log
so one orchestrator run (and its child scripts) can be traced in one place.

    from log_config import setup_logging
    log = setup_logging("approval_report_data")
"""
import sys
import logging
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).resolve().parent
LOG_DIR = BASE_DIR / "logs"
LOG_PREFIX = "approved_facility"
LOG_FORMAT = "%(asctime)s | %(levelname)-7s | %(name)-24s | %(message)s"
DATE_FORMAT = "%Y-%m-%d %H:%M:%S"


def log_file_path(day=None):
    return LOG_DIR / f"{LOG_PREFIX}_{(day or datetime.now()).strftime('%Y-%m-%d')}.log"


def setup_logging(script_name, level=logging.INFO):
    """Configure the root logger once (console + daily file); return a logger for the script."""
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    root = logging.getLogger()
    if not getattr(root, "_appr_af_configured", False):
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        formatter = logging.Formatter(LOG_FORMAT, DATE_FORMAT)

        console = logging.StreamHandler(sys.stdout)
        console.setFormatter(formatter)
        file_handler = logging.FileHandler(log_file_path(), encoding="utf-8")
        file_handler.setFormatter(formatter)

        root.handlers.clear()
        root.addHandler(console)
        root.addHandler(file_handler)
        root.setLevel(level)
        # Third-party noise
        for noisy in ("selenium", "urllib3", "PIL", "matplotlib", "fontTools"):
            logging.getLogger(noisy).setLevel(logging.WARNING)
        root._appr_af_configured = True

        def _excepthook(exc_type, exc, tb):
            if issubclass(exc_type, KeyboardInterrupt):
                sys.__excepthook__(exc_type, exc, tb)
                return
            logging.getLogger(script_name).critical("Unhandled exception",
                                                    exc_info=(exc_type, exc, tb))
        sys.excepthook = _excepthook

    log = logging.getLogger(script_name)
    log.info(f"===== {script_name} started | args: {' '.join(sys.argv[1:]) or '-'} "
             f"| log: {log_file_path().name} =====")
    return log
