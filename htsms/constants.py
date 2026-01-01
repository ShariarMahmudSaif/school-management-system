from __future__ import annotations

from pathlib import Path

APP_NAME = "High Tech School Management System"

WORKSPACE_ROOT = Path(__file__).resolve().parents[1]
DATA_XLSX_PATH = WORKSPACE_ROOT / "school_data.xlsx"
SETTINGS_JSON_PATH = WORKSPACE_ROOT / "settings.json"
ERROR_LOG_PATH = WORKSPACE_ROOT / "error_log.txt"

STUDENTS_SHEET = "students"
TEACHERS_SHEET = "teachers"
STUDENT_PAYMENTS_SHEET = "student_payments"
TEACHER_PAYMENTS_SHEET = "teacher_payments"
ACTIVITY_SHEET = "activity_log"
