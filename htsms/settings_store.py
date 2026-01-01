from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from .constants import SETTINGS_JSON_PATH


@dataclass
class Settings:
    student_id_prefix: str = "STU-"
    teacher_id_prefix: str = "TCH-"
    student_custom_fields: list[str] = field(default_factory=list)
    teacher_custom_fields: list[str] = field(default_factory=list)
    default_year: int = 2026
    default_month: int = 1
    appearance_mode: str = "Light"  # Light | Dark | System
    ui_scaling: float = 1.0
    default_student_fee: float = 0.0  # Monthly tuition fee
    default_teacher_salary: float = 0.0  # Monthly salary

    @staticmethod
    def from_dict(d: dict[str, Any]) -> "Settings":
        raw_scale = d.get("ui_scaling", 1.0)
        try:
            scale = float(raw_scale)
        except Exception:
            scale = 1.0
        # Keep scaling in a sane range to avoid blurry fractional scaling.
        if scale < 0.8:
            scale = 0.8
        if scale > 1.4:
            scale = 1.4
        try:
            student_fee = float(d.get("default_student_fee", 0.0))
        except Exception:
            student_fee = 0.0
        try:
            teacher_salary = float(d.get("default_teacher_salary", 0.0))
        except Exception:
            teacher_salary = 0.0
        return Settings(
            student_id_prefix=str(d.get("student_id_prefix", "STU-")),
            teacher_id_prefix=str(d.get("teacher_id_prefix", "TCH-")),
            student_custom_fields=list(d.get("student_custom_fields", [])),
            teacher_custom_fields=list(d.get("teacher_custom_fields", [])),
            default_year=int(d.get("default_year", 2026)),
            default_month=int(d.get("default_month", 1)),
            appearance_mode=str(d.get("appearance_mode", "Light")),
            ui_scaling=scale,
            default_student_fee=student_fee,
            default_teacher_salary=teacher_salary,
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "student_id_prefix": self.student_id_prefix,
            "teacher_id_prefix": self.teacher_id_prefix,
            "student_custom_fields": self.student_custom_fields,
            "teacher_custom_fields": self.teacher_custom_fields,
            "default_year": self.default_year,
            "default_month": self.default_month,
            "appearance_mode": self.appearance_mode,
            "ui_scaling": self.ui_scaling,
            "default_student_fee": self.default_student_fee,
            "default_teacher_salary": self.default_teacher_salary,
        }


class SettingsStore:
    def __init__(self, path: Path = SETTINGS_JSON_PATH):
        self.path = path

    def load(self) -> Settings:
        if not self.path.exists():
            settings = Settings()
            self.save(settings)
            return settings

        with self.path.open("r", encoding="utf-8") as f:
            data = json.load(f)
        return Settings.from_dict(data if isinstance(data, dict) else {})

    def save(self, settings: Settings) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self.path.open("w", encoding="utf-8") as f:
            json.dump(settings.to_dict(), f, indent=2)
            f.write("\n")
