"""Structured logger for the Auto-PPT Agent."""

import logging
import sys
from datetime import datetime


def get_logger(name: str = "auto_ppt") -> logging.Logger:
    logger = logging.getLogger(name)
    if not logger.handlers:
        handler = logging.StreamHandler(sys.stdout)
        fmt = logging.Formatter(
            "[%(asctime)s] %(levelname)-8s %(name)s – %(message)s",
            datefmt="%H:%M:%S",
        )
        handler.setFormatter(fmt)
        logger.addHandler(handler)
        logger.setLevel(logging.DEBUG)
    return logger


# ── Step-level helpers used by Streamlit ──────────────────────────────────────

class StepLog:
    """Collects log entries for display in the Streamlit UI."""

    def __init__(self):
        self.entries: list[dict] = []

    def add(self, step: str, status: str, detail: str = ""):
        self.entries.append({
            "time": datetime.now().strftime("%H:%M:%S"),
            "step": step,
            "status": status,   # "ok" | "error" | "info"
            "detail": detail,
        })

    def ok(self, step: str, detail: str = ""):
        self.add(step, "ok", detail)

    def error(self, step: str, detail: str = ""):
        self.add(step, "error", detail)

    def info(self, step: str, detail: str = ""):
        self.add(step, "info", detail)