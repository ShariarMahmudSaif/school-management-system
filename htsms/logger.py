from __future__ import annotations

import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from .constants import ERROR_LOG_PATH


@dataclass
class AppEvent:
    timestamp: str
    action: str
    entity_type: str
    entity_id: str
    details: str = ""


class ErrorLogger:
    def __init__(self, path: Path = ERROR_LOG_PATH):
        self.path = path

    def log_exception(self, exc: BaseException, context: str = "") -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().isoformat(timespec="seconds")
        with self.path.open("a", encoding="utf-8") as f:
            f.write(f"[{ts}] {context}\n")
            f.write("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
            f.write("\n")


def now_ts() -> str:
    return datetime.now().isoformat(timespec="seconds")
