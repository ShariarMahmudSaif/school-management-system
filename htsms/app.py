from __future__ import annotations

import ctypes
import re
import sys
from dataclasses import dataclass
from tkinter import ttk
from typing import Any

import customtkinter as ctk

from .constants import APP_NAME
from .logger import AppEvent, ErrorLogger, now_ts
from .settings_store import SettingsStore
from .storage import ExcelStore


MONTHS = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
]


def _safe_int(s: str, default: int) -> int:
    try:
        return int(s)
    except Exception:
        return default


def _normalize_field_name(name: str) -> str:
    name = name.strip()
    name = re.sub(r"\s+", "_", name)
    name = re.sub(r"[^a-zA-Z0-9_\-]", "", name)
    return name


def _enable_high_dpi() -> None:
    """Best-effort: make the app crisp on Windows high-DPI displays."""

    if not sys.platform.startswith("win"):
        return
    try:
        # Windows 8.1+
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # type: ignore[attr-defined]
    except Exception:
        try:
            # Older fallback
            ctypes.windll.user32.SetProcessDPIAware()  # type: ignore[attr-defined]
        except Exception:
            pass


@dataclass
class Selected:
    entity: str = ""
    person_id: str = ""


class HTSMSApp(ctk.CTk):
    """High Tech School Management System UI.

    Notes on UI goals:
    - Crisp rendering: high-DPI awareness + sane scaling.
    - Fast lists: ttk.Treeview (much faster than hundreds of buttons).
    - "Web-like" layout: sidebar navigation + card panels.
    """

    def __init__(self):
        super().__init__()

        self.err_logger = ErrorLogger()
        self.settings_store = SettingsStore()
        self.settings = self.settings_store.load()

        # Theme / scaling first (before building widgets)
        ctk.set_default_color_theme("blue")
        self._apply_ui_settings()

        self.title(APP_NAME)
        self.geometry("1240x760")
        self.minsize(1080, 680)

        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self.store = ExcelStore()
        self.store.ensure_workbook(self.settings.student_custom_fields, self.settings.teacher_custom_fields)

        self.selected = Selected()
        self._after_ids: dict[str, str] = {}
        self._nav_buttons: dict[str, ctk.CTkButton] = {}
        self._pages: dict[str, ctk.CTkFrame] = {}

        self._configure_ttk()
        self._build_shell()
        self.refresh_all()
        self.show_page("dashboard")

    # Tkinter callback errors can be silent; log them.
    def report_callback_exception(self, exc, val, tb):  # type: ignore[override]
        try:
            self.err_logger.log_exception(val, f"tk_callback: {exc}")
        finally:
            super().report_callback_exception(exc, val, tb)

    # ---------------- Look & Feel ----------------
    def _apply_ui_settings(self) -> None:
        mode = (self.settings.appearance_mode or "System").strip().capitalize()
        if mode not in {"Light", "Dark", "System"}:
            mode = "System"
        ctk.set_appearance_mode(mode)

        try:
            scale = float(getattr(self.settings, "ui_scaling", 1.0) or 1.0)
        except Exception:
            scale = 1.0
        if scale < 0.8:
            scale = 0.8
        if scale > 1.4:
            scale = 1.4
        # Using the same value for window+widget scaling avoids fractional blur.
        ctk.set_window_scaling(scale)
        ctk.set_widget_scaling(scale)

    def _configure_ttk(self) -> None:
        # Make Treeview look "clean" and readable.
        try:
            style = ttk.Style()
            style.theme_use("clam")
            style.configure(
                "Treeview",
                rowheight=28,
                borderwidth=0,
                relief="flat",
            )
            style.configure(
                "Treeview.Heading",
                font=("Segoe UI", 10, "bold"),
                padding=(8, 6),
            )
        except Exception as e:
            self.err_logger.log_exception(e, "configure_ttk")

    # ---------------- Shell / navigation ----------------
    def _build_shell(self) -> None:
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.sidebar = ctk.CTkFrame(self, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsw")
        self.sidebar.grid_rowconfigure(10, weight=1)

        brand = ctk.CTkFrame(self.sidebar, corner_radius=0)
        brand.grid(row=0, column=0, sticky="ew", padx=14, pady=(18, 10))
        brand.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(brand, text="High Tech", font=ctk.CTkFont(size=18, weight="bold")).grid(
            row=0, column=0, sticky="w"
        )
        ctk.CTkLabel(brand, text="School Management", font=ctk.CTkFont(size=12)).grid(row=1, column=0, sticky="w")

        self._nav_item("dashboard", "Dashboard", 1)
        self._nav_item("students", "Students", 2)
        self._nav_item("teachers", "Teachers", 3)
        self._nav_item("activity", "Activity Log", 4)
        self._nav_item("settings", "Settings", 5)

        ctk.CTkLabel(
            self.sidebar,
            text="Data: school_data.xlsx\nLogs: error_log.txt",
            justify="left",
            font=ctk.CTkFont(size=11),
            text_color=("#6b7280", "#94a3b8"),
        ).grid(row=11, column=0, sticky="sw", padx=14, pady=(10, 14))

        self.content = ctk.CTkFrame(self, corner_radius=0)
        self.content.grid(row=0, column=1, sticky="nsew")
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(0, weight=1)

        # Pages stack
        self.pages_stack = ctk.CTkFrame(self.content, corner_radius=0)
        self.pages_stack.grid(row=0, column=0, sticky="nsew")
        self.pages_stack.grid_columnconfigure(0, weight=1)
        self.pages_stack.grid_rowconfigure(0, weight=1)

        self._pages["dashboard"] = self._build_dashboard_page(self.pages_stack)
        self._pages["students"] = self._build_students_page(self.pages_stack)
        self._pages["teachers"] = self._build_teachers_page(self.pages_stack)
        self._pages["activity"] = self._build_activity_page(self.pages_stack)
        self._pages["settings"] = self._build_settings_page(self.pages_stack)

        for p in self._pages.values():
            p.grid(row=0, column=0, sticky="nsew")
            p.grid_remove()

    def _nav_item(self, key: str, label: str, row: int) -> None:
        btn = ctk.CTkButton(
            self.sidebar,
            text=label,
            anchor="w",
            height=40,
            corner_radius=10,
            fg_color="transparent",
            hover_color=("#e5e7eb", "#1f2937"),
            command=lambda k=key: self.show_page(k),
        )
        btn.grid(row=row, column=0, sticky="ew", padx=14, pady=6)
        self._nav_buttons[key] = btn

    def show_page(self, key: str) -> None:
        for k, page in self._pages.items():
            if k == key:
                page.grid()
            else:
                page.grid_remove()

        for k, btn in self._nav_buttons.items():
            if k == key:
                btn.configure(fg_color=("#dbeafe", "#0b3b70"), text_color=("#111827", "#ffffff"))
            else:
                btn.configure(fg_color="transparent", text_color=("#111827", "#ffffff"))

        # Keep pages feeling snappy by refreshing only what matters.
        if key == "dashboard":
            self.refresh_dashboard()
        elif key == "students":
            self.refresh_students()
        elif key == "teachers":
            self.refresh_teachers()
        elif key == "activity":
            self.refresh_activity()
        elif key == "settings":
            self.refresh_custom_fields_views()

    # ---------------- Reusable "card" helpers ----------------
    def _card(self, parent: ctk.CTkFrame, **kwargs: Any) -> ctk.CTkFrame:
        return ctk.CTkFrame(
            parent,
            corner_radius=14,
            border_width=1,
            border_color=("#e5e7eb", "#1f2937"),
            **kwargs,
        )

    def _page_title(self, parent: ctk.CTkFrame, title: str, subtitle: str) -> ctk.CTkFrame:
        bar = ctk.CTkFrame(parent, corner_radius=0)
        bar.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(bar, text=title, font=ctk.CTkFont(size=22, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=18, pady=(18, 2)
        )
        ctk.CTkLabel(bar, text=subtitle, font=ctk.CTkFont(size=12), text_color=("#6b7280", "#94a3b8")).grid(
            row=1, column=0, sticky="w", padx=18, pady=(0, 12)
        )
        return bar

    def _debounce(self, key: str, delay_ms: int, fn) -> None:
        if key in self._after_ids:
            try:
                self.after_cancel(self._after_ids[key])
            except Exception:
                pass
        self._after_ids[key] = self.after(delay_ms, fn)

    # ---------------- Pages ----------------
    def _build_dashboard_page(self, parent: ctk.CTkFrame) -> ctk.CTkFrame:
        page = ctk.CTkFrame(parent, corner_radius=0)
        page.grid_rowconfigure(2, weight=1)
        page.grid_columnconfigure(0, weight=1)

        self._page_title(page, "Dashboard", "Quick overview of students, teachers, and payments").grid(row=0, column=0, sticky="ew")

        filters = self._card(page)
        filters.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 12))
        filters.grid_columnconfigure(5, weight=1)

        ctk.CTkLabel(filters, text="Month").grid(row=0, column=0, padx=(12, 6), pady=12)
        self.dash_month = ctk.CTkOptionMenu(filters, values=MONTHS)
        self.dash_month.grid(row=0, column=1, padx=6, pady=12)

        ctk.CTkLabel(filters, text="Year").grid(row=0, column=2, padx=(12, 6), pady=12)
        self.dash_year = ctk.CTkEntry(filters, width=120)
        self.dash_year.grid(row=0, column=3, padx=6, pady=12)

        ctk.CTkButton(filters, text="Refresh", command=self.refresh_dashboard).grid(row=0, column=4, padx=12, pady=12)

        body = ctk.CTkFrame(page, corner_radius=0)
        body.grid(row=2, column=0, sticky="nsew", padx=18, pady=(0, 18))
        body.grid_columnconfigure((0, 1), weight=1)
        body.grid_rowconfigure(1, weight=1)

        self.card_students = self._card(body)
        self.card_students.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=(0, 10))
        self.card_teachers = self._card(body)
        self.card_teachers.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=(0, 10))

        self.dash_students = ctk.CTkLabel(self.card_students, text="Students", font=ctk.CTkFont(size=16, weight="bold"))
        self.dash_students.grid(row=0, column=0, sticky="w", padx=16, pady=(16, 6))
        self.dash_students_paid = ctk.CTkLabel(self.card_students, text="Tuition (Paid/Pending): ...")
        self.dash_students_paid.grid(row=1, column=0, sticky="w", padx=16, pady=(0, 16))

        self.dash_teachers = ctk.CTkLabel(self.card_teachers, text="Teachers", font=ctk.CTkFont(size=16, weight="bold"))
        self.dash_teachers.grid(row=0, column=0, sticky="w", padx=16, pady=(16, 6))
        self.dash_teachers_paid = ctk.CTkLabel(self.card_teachers, text="Salary (Paid/Pending): ...")
        self.dash_teachers_paid.grid(row=1, column=0, sticky="w", padx=16, pady=(0, 16))

        hint = self._card(body)
        hint.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=0, pady=(10, 0))
        ctk.CTkLabel(
            hint,
            text=(
                "Payments are tracked per month/year. Go to Students/Teachers to toggle Paid/Pending for the selected period.\n"
                "Tip: Use search to quickly find IDs, names, classes, sections, or contacts."
            ),
            justify="left",
            wraplength=980,
            text_color=("#374151", "#cbd5e1"),
        ).grid(row=0, column=0, sticky="w", padx=16, pady=16)

        return page

    def _build_students_page(self, parent: ctk.CTkFrame) -> ctk.CTkFrame:
        page = ctk.CTkFrame(parent, corner_radius=0)
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(2, weight=1)

        self._page_title(page, "Students", "Profiles, contacts, and tuition payments").grid(row=0, column=0, sticky="ew")

        toolbar = self._card(page)
        toolbar.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 12))
        toolbar.grid_columnconfigure(6, weight=1)

        ctk.CTkLabel(toolbar, text="Search").grid(row=0, column=0, padx=(12, 6), pady=12)
        self.student_search = ctk.CTkEntry(toolbar, placeholder_text="name, ID, class, section, contact")
        self.student_search.grid(row=0, column=1, padx=6, pady=12, sticky="ew")
        self.student_search.bind("<KeyRelease>", lambda _e: self._debounce("student_search", 180, self.refresh_students))

        ctk.CTkLabel(toolbar, text="Class").grid(row=0, column=2, padx=(12, 6), pady=12)
        self.student_class_filter = ctk.CTkEntry(toolbar, placeholder_text="e.g. 10", width=120)
        self.student_class_filter.grid(row=0, column=3, padx=6, pady=12)
        self.student_class_filter.bind("<KeyRelease>", lambda _e: self._debounce("student_filter", 180, self.refresh_students))

        ctk.CTkButton(toolbar, text="Add", command=self.add_student).grid(row=0, column=4, padx=(12, 6), pady=12)
        ctk.CTkButton(toolbar, text="Refresh", command=self.refresh_students).grid(row=0, column=5, padx=6, pady=12)

        body = ctk.CTkFrame(page, corner_radius=0)
        body.grid(row=2, column=0, sticky="nsew", padx=18, pady=(0, 18))
        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=0)
        body.grid_rowconfigure(0, weight=1)

        left = self._card(body)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12), pady=0)
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(left, corner_radius=0)
        header.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 6))
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text="Students", font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, sticky="w")
        self.students_count = ctk.CTkLabel(header, text="0 records", text_color=("#6b7280", "#94a3b8"))
        self.students_count.grid(row=0, column=1, sticky="e")

        self.students_tree = self._make_students_tree(left)
        self.students_tree.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        self.students_tree.bind("<<TreeviewSelect>>", self._on_student_selected)

        right = ctk.CTkFrame(body, corner_radius=0)
        right.grid(row=0, column=1, sticky="ns", padx=(12, 0), pady=0)
        right.grid_rowconfigure(10, weight=1)

        actions = self._card(right)
        actions.grid(row=0, column=0, sticky="ew", padx=0, pady=(0, 12))
        actions.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(actions, text="Actions", font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=14, pady=(14, 6)
        )
        ctk.CTkButton(actions, text="Edit Selected", command=self.edit_selected_student).grid(row=1, column=0, padx=14, pady=6, sticky="ew")
        ctk.CTkButton(
            actions,
            text="Delete Selected",
            fg_color="#b91c1c",
            hover_color="#991b1b",
            command=self.delete_selected_student,
        ).grid(row=2, column=0, padx=14, pady=(6, 14), sticky="ew")

        pay = self._card(right)
        pay.grid(row=1, column=0, sticky="ew")
        pay.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(pay, text="Tuition Payment", font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=14, pady=(14, 6)
        )
        ctk.CTkLabel(pay, text="Month").grid(row=1, column=0, padx=14, pady=8, sticky="w")
        self.student_month = ctk.CTkOptionMenu(pay, values=MONTHS, command=lambda _v: self._update_student_payment_label())
        self.student_month.grid(row=1, column=1, padx=14, pady=8, sticky="ew")

        ctk.CTkLabel(pay, text="Year").grid(row=2, column=0, padx=14, pady=8, sticky="w")
        self.student_year = ctk.CTkEntry(pay)
        self.student_year.grid(row=2, column=1, padx=14, pady=8, sticky="ew")
        self.student_year.bind("<KeyRelease>", lambda _e: self._debounce("student_year", 180, self._update_student_payment_label))

        self.student_payment_status = ctk.CTkLabel(pay, text="Status: (select a student)")
        self.student_payment_status.grid(row=3, column=0, columnspan=2, padx=14, pady=(10, 6), sticky="w")
        ctk.CTkButton(pay, text="Toggle Paid / Pending", command=self.toggle_student_payment).grid(
            row=4, column=0, columnspan=2, padx=14, pady=(6, 14), sticky="ew"
        )

        return page

    def _build_teachers_page(self, parent: ctk.CTkFrame) -> ctk.CTkFrame:
        page = ctk.CTkFrame(parent, corner_radius=0)
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(2, weight=1)

        self._page_title(page, "Teachers", "Profiles, contacts, and salary payments").grid(row=0, column=0, sticky="ew")

        toolbar = self._card(page)
        toolbar.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 12))
        toolbar.grid_columnconfigure(5, weight=1)

        ctk.CTkLabel(toolbar, text="Search").grid(row=0, column=0, padx=(12, 6), pady=12)
        self.teacher_search = ctk.CTkEntry(toolbar, placeholder_text="name, ID, role, contact")
        self.teacher_search.grid(row=0, column=1, padx=6, pady=12, sticky="ew")
        self.teacher_search.bind("<KeyRelease>", lambda _e: self._debounce("teacher_search", 180, self.refresh_teachers))

        ctk.CTkButton(toolbar, text="Add", command=self.add_teacher).grid(row=0, column=2, padx=(12, 6), pady=12)
        ctk.CTkButton(toolbar, text="Refresh", command=self.refresh_teachers).grid(row=0, column=3, padx=6, pady=12)

        body = ctk.CTkFrame(page, corner_radius=0)
        body.grid(row=2, column=0, sticky="nsew", padx=18, pady=(0, 18))
        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=0)
        body.grid_rowconfigure(0, weight=1)

        left = self._card(body)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12), pady=0)
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(left, corner_radius=0)
        header.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 6))
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text="Teachers", font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, sticky="w")
        self.teachers_count = ctk.CTkLabel(header, text="0 records", text_color=("#6b7280", "#94a3b8"))
        self.teachers_count.grid(row=0, column=1, sticky="e")

        self.teachers_tree = self._make_teachers_tree(left)
        self.teachers_tree.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        self.teachers_tree.bind("<<TreeviewSelect>>", self._on_teacher_selected)

        right = ctk.CTkFrame(body, corner_radius=0)
        right.grid(row=0, column=1, sticky="ns", padx=(12, 0), pady=0)
        right.grid_rowconfigure(10, weight=1)

        actions = self._card(right)
        actions.grid(row=0, column=0, sticky="ew", padx=0, pady=(0, 12))
        actions.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(actions, text="Actions", font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=14, pady=(14, 6)
        )
        ctk.CTkButton(actions, text="Edit Selected", command=self.edit_selected_teacher).grid(row=1, column=0, padx=14, pady=6, sticky="ew")
        ctk.CTkButton(
            actions,
            text="Delete Selected",
            fg_color="#b91c1c",
            hover_color="#991b1b",
            command=self.delete_selected_teacher,
        ).grid(row=2, column=0, padx=14, pady=(6, 14), sticky="ew")

        pay = self._card(right)
        pay.grid(row=1, column=0, sticky="ew")
        pay.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(pay, text="Salary Payment", font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=14, pady=(14, 6)
        )
        ctk.CTkLabel(pay, text="Month").grid(row=1, column=0, padx=14, pady=8, sticky="w")
        self.teacher_month = ctk.CTkOptionMenu(pay, values=MONTHS, command=lambda _v: self._update_teacher_payment_label())
        self.teacher_month.grid(row=1, column=1, padx=14, pady=8, sticky="ew")

        ctk.CTkLabel(pay, text="Year").grid(row=2, column=0, padx=14, pady=8, sticky="w")
        self.teacher_year = ctk.CTkEntry(pay)
        self.teacher_year.grid(row=2, column=1, padx=14, pady=8, sticky="ew")
        self.teacher_year.bind("<KeyRelease>", lambda _e: self._debounce("teacher_year", 180, self._update_teacher_payment_label))

        self.teacher_payment_status = ctk.CTkLabel(pay, text="Status: (select a teacher)")
        self.teacher_payment_status.grid(row=3, column=0, columnspan=2, padx=14, pady=(10, 6), sticky="w")
        ctk.CTkButton(pay, text="Toggle Paid / Pending", command=self.toggle_teacher_payment).grid(
            row=4, column=0, columnspan=2, padx=14, pady=(6, 14), sticky="ew"
        )

        return page

    def _build_activity_page(self, parent: ctk.CTkFrame) -> ctk.CTkFrame:
        page = ctk.CTkFrame(parent, corner_radius=0)
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(2, weight=1)

        self._page_title(page, "Activity Log", "Everything that happens is recorded here").grid(row=0, column=0, sticky="ew")

        toolbar = self._card(page)
        toolbar.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 12))
        toolbar.grid_columnconfigure(3, weight=1)
        ctk.CTkLabel(toolbar, text="Search").grid(row=0, column=0, padx=(12, 6), pady=12)
        self.activity_search = ctk.CTkEntry(toolbar, placeholder_text="action, entity, details")
        self.activity_search.grid(row=0, column=1, padx=6, pady=12, sticky="ew")
        self.activity_search.bind("<KeyRelease>", lambda _e: self._debounce("activity_search", 180, self.refresh_activity))
        ctk.CTkButton(toolbar, text="Refresh", command=self.refresh_activity).grid(row=0, column=2, padx=12, pady=12)

        body = self._card(page)
        body.grid(row=2, column=0, sticky="nsew", padx=18, pady=(0, 18))
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(0, weight=1)

        self.activity_tree = self._make_activity_tree(body)
        self.activity_tree.grid(row=0, column=0, sticky="nsew", padx=12, pady=12)
        return page

    def _build_settings_page(self, parent: ctk.CTkFrame) -> ctk.CTkFrame:
        page = ctk.CTkFrame(parent, corner_radius=0)
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(2, weight=1)

        self._page_title(page, "Settings", "Customize prefixes, fields, and UI preferences").grid(row=0, column=0, sticky="ew")

        top = self._card(page)
        top.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 12))
        top.grid_columnconfigure((1, 3, 5), weight=1)

        ctk.CTkLabel(top, text="Student ID Prefix").grid(row=0, column=0, padx=(12, 6), pady=12, sticky="w")
        self.student_prefix = ctk.CTkEntry(top)
        self.student_prefix.grid(row=0, column=1, padx=6, pady=12, sticky="ew")

        ctk.CTkLabel(top, text="Teacher ID Prefix").grid(row=0, column=2, padx=(12, 6), pady=12, sticky="w")
        self.teacher_prefix = ctk.CTkEntry(top)
        self.teacher_prefix.grid(row=0, column=3, padx=6, pady=12, sticky="ew")

        ctk.CTkLabel(top, text="Appearance").grid(row=0, column=4, padx=(12, 6), pady=12, sticky="w")
        self.appearance_mode = ctk.CTkOptionMenu(top, values=["Light", "Dark", "System"])
        self.appearance_mode.grid(row=0, column=5, padx=6, pady=12, sticky="ew")

        ctk.CTkLabel(top, text="UI Scale").grid(row=1, column=0, padx=(12, 6), pady=12, sticky="w")
        self.ui_scale = ctk.CTkOptionMenu(top, values=["0.9", "1.0", "1.1", "1.2", "1.3"])
        self.ui_scale.grid(row=1, column=1, padx=6, pady=12, sticky="w")

        ctk.CTkButton(top, text="Save Settings", command=self.save_settings).grid(row=1, column=5, padx=12, pady=12, sticky="e")

        body = ctk.CTkFrame(page, corner_radius=0)
        body.grid(row=2, column=0, sticky="nsew", padx=18, pady=(0, 18))
        body.grid_columnconfigure((0, 1), weight=1)
        body.grid_rowconfigure(0, weight=1)

        s_box = self._card(body)
        s_box.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=0)
        s_box.grid_columnconfigure(0, weight=1)
        s_box.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(s_box, text="Student Custom Fields", font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, padx=14, pady=(14, 6), sticky="w"
        )
        row = ctk.CTkFrame(s_box, corner_radius=0)
        row.grid(row=1, column=0, sticky="ew", padx=14, pady=6)
        row.grid_columnconfigure(0, weight=1)
        self.student_cf_entry = ctk.CTkEntry(row, placeholder_text="e.g. guardian_name")
        self.student_cf_entry.grid(row=0, column=0, padx=(0, 8), pady=0, sticky="ew")
        ctk.CTkButton(row, text="Add", width=90, command=self.add_student_custom_field).grid(row=0, column=1)

        self.student_cf_list = ctk.CTkScrollableFrame(s_box)
        self.student_cf_list.grid(row=2, column=0, padx=14, pady=(6, 14), sticky="nsew")
        self.student_cf_list.grid_columnconfigure(0, weight=1)

        t_box = self._card(body)
        t_box.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=0)
        t_box.grid_columnconfigure(0, weight=1)
        t_box.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(t_box, text="Teacher Custom Fields", font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, padx=14, pady=(14, 6), sticky="w"
        )
        row = ctk.CTkFrame(t_box, corner_radius=0)
        row.grid(row=1, column=0, sticky="ew", padx=14, pady=6)
        row.grid_columnconfigure(0, weight=1)
        self.teacher_cf_entry = ctk.CTkEntry(row, placeholder_text="e.g. qualification")
        self.teacher_cf_entry.grid(row=0, column=0, padx=(0, 8), pady=0, sticky="ew")
        ctk.CTkButton(row, text="Add", width=90, command=self.add_teacher_custom_field).grid(row=0, column=1)

        self.teacher_cf_list = ctk.CTkScrollableFrame(t_box)
        self.teacher_cf_list.grid(row=2, column=0, padx=14, pady=(6, 14), sticky="nsew")
        self.teacher_cf_list.grid_columnconfigure(0, weight=1)

        return page

    # ---------------- Treeviews ----------------
    def _wrap_tree(self, parent: ctk.CTkFrame) -> ctk.CTkFrame:
        wrap = ctk.CTkFrame(parent, corner_radius=0)
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_rowconfigure(0, weight=1)
        return wrap

    def _make_students_tree(self, parent: ctk.CTkFrame) -> ttk.Treeview:
        wrap = self._wrap_tree(parent)
        columns = ("id", "name", "class", "section", "primary", "secondary")
        tree = ttk.Treeview(wrap, columns=columns, show="headings", selectmode="browse")
        tree.heading("id", text="ID")
        tree.heading("name", text="Name")
        tree.heading("class", text="Class")
        tree.heading("section", text="Section")
        tree.heading("primary", text="Primary Contact")
        tree.heading("secondary", text="Secondary Contact")
        tree.column("id", width=120, anchor="w")
        tree.column("name", width=220, anchor="w")
        tree.column("class", width=80, anchor="w")
        tree.column("section", width=80, anchor="w")
        tree.column("primary", width=180, anchor="w")
        tree.column("secondary", width=180, anchor="w")
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        wrap.grid(row=1, column=0, sticky="nsew")
        return tree

    def _make_teachers_tree(self, parent: ctk.CTkFrame) -> ttk.Treeview:
        wrap = self._wrap_tree(parent)
        columns = ("id", "name", "role", "primary", "secondary")
        tree = ttk.Treeview(wrap, columns=columns, show="headings", selectmode="browse")
        tree.heading("id", text="ID")
        tree.heading("name", text="Name")
        tree.heading("role", text="Role")
        tree.heading("primary", text="Primary Contact")
        tree.heading("secondary", text="Secondary Contact")
        tree.column("id", width=120, anchor="w")
        tree.column("name", width=240, anchor="w")
        tree.column("role", width=180, anchor="w")
        tree.column("primary", width=180, anchor="w")
        tree.column("secondary", width=180, anchor="w")
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        wrap.grid(row=1, column=0, sticky="nsew")
        return tree

    def _make_activity_tree(self, parent: ctk.CTkFrame) -> ttk.Treeview:
        wrap = self._wrap_tree(parent)
        columns = ("ts", "action", "entity", "details")
        tree = ttk.Treeview(wrap, columns=columns, show="headings", selectmode="browse")
        tree.heading("ts", text="Timestamp")
        tree.heading("action", text="Action")
        tree.heading("entity", text="Entity")
        tree.heading("details", text="Details")
        tree.column("ts", width=170, anchor="w")
        tree.column("action", width=140, anchor="w")
        tree.column("entity", width=160, anchor="w")
        tree.column("details", width=600, anchor="w")
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        wrap.grid(row=0, column=0, sticky="nsew")
        return tree

    # ---------------- Data refresh ----------------
    def refresh_all(self) -> None:
        # Month / year defaults
        self.dash_month.set(MONTHS[self.settings.default_month - 1])
        self.dash_year.delete(0, "end")
        self.dash_year.insert(0, str(self.settings.default_year))

        self.student_month.set(MONTHS[self.settings.default_month - 1])
        self.student_year.delete(0, "end")
        self.student_year.insert(0, str(self.settings.default_year))

        self.teacher_month.set(MONTHS[self.settings.default_month - 1])
        self.teacher_year.delete(0, "end")
        self.teacher_year.insert(0, str(self.settings.default_year))

        # Settings values
        self.student_prefix.delete(0, "end")
        self.student_prefix.insert(0, self.settings.student_id_prefix)
        self.teacher_prefix.delete(0, "end")
        self.teacher_prefix.insert(0, self.settings.teacher_id_prefix)
        self.appearance_mode.set((self.settings.appearance_mode or "System").capitalize())
        self.ui_scale.set(str(getattr(self.settings, "ui_scaling", 1.0)))

        self.refresh_students()
        self.refresh_teachers()
        self.refresh_dashboard()
        self.refresh_activity()
        self.refresh_custom_fields_views()

    def _month_index(self, month_name: str) -> int:
        try:
            return MONTHS.index(month_name) + 1
        except ValueError:
            return self.settings.default_month

    def _emit(self, action: str, entity_type: str, entity_id: str, details: str = "") -> None:
        self.store.add_event(AppEvent(timestamp=now_ts(), action=action, entity_type=entity_type, entity_id=entity_id, details=details))

    def refresh_students(self) -> None:
        try:
            data = self.store.list_students()
            q = (self.student_search.get() or "").strip().lower()
            class_filter = (self.student_class_filter.get() or "").strip().lower()

            filtered = []
            for r in data:
                if class_filter and str(r.get("class", "")).strip().lower() != class_filter:
                    continue
                if q:
                    blob = " ".join(
                        [
                            str(r.get("student_id", "")),
                            str(r.get("first_name", "")),
                            str(r.get("last_name", "")),
                            str(r.get("class", "")),
                            str(r.get("section", "")),
                            str(r.get("primary_contact", "")),
                            str(r.get("secondary_contact", "")),
                        ]
                    ).lower()
                    if q not in blob:
                        continue
                filtered.append(r)

            filtered.sort(key=lambda r: str(r.get("student_id", "")))

            for item in self.students_tree.get_children(""):
                self.students_tree.delete(item)

            for r in filtered:
                sid = str(r.get("student_id", ""))
                name = f"{r.get('first_name','')} {r.get('last_name','')}".strip()
                cls = str(r.get("class", ""))
                sec = str(r.get("section", ""))
                primary = str(r.get("primary_contact", ""))
                secondary = str(r.get("secondary_contact", ""))
                if not sid:
                    continue
                self.students_tree.insert("", "end", iid=sid, values=(sid, name, cls, sec, primary, secondary))

            self.students_count.configure(text=f"{len(filtered)} records")

            if self.selected.entity == "student" and self.selected.person_id:
                if self.selected.person_id in self.students_tree.get_children(""):
                    self.students_tree.selection_set(self.selected.person_id)
                self._update_student_payment_label()
        except Exception as e:
            self.err_logger.log_exception(e, "refresh_students")

    def refresh_teachers(self) -> None:
        try:
            data = self.store.list_teachers()
            q = (self.teacher_search.get() or "").strip().lower()
            filtered = []
            for r in data:
                if q:
                    blob = " ".join(
                        [
                            str(r.get("teacher_id", "")),
                            str(r.get("first_name", "")),
                            str(r.get("last_name", "")),
                            str(r.get("role", "")),
                            str(r.get("primary_contact", "")),
                            str(r.get("secondary_contact", "")),
                        ]
                    ).lower()
                    if q not in blob:
                        continue
                filtered.append(r)

            filtered.sort(key=lambda r: str(r.get("teacher_id", "")))

            for item in self.teachers_tree.get_children(""):
                self.teachers_tree.delete(item)

            for r in filtered:
                tid = str(r.get("teacher_id", ""))
                name = f"{r.get('first_name','')} {r.get('last_name','')}".strip()
                role = str(r.get("role", ""))
                primary = str(r.get("primary_contact", ""))
                secondary = str(r.get("secondary_contact", ""))
                if not tid:
                    continue
                self.teachers_tree.insert("", "end", iid=tid, values=(tid, name, role, primary, secondary))

            self.teachers_count.configure(text=f"{len(filtered)} records")

            if self.selected.entity == "teacher" and self.selected.person_id:
                if self.selected.person_id in self.teachers_tree.get_children(""):
                    self.teachers_tree.selection_set(self.selected.person_id)
                self._update_teacher_payment_label()
        except Exception as e:
            self.err_logger.log_exception(e, "refresh_teachers")

    def refresh_activity(self) -> None:
        try:
            q = (self.activity_search.get() or "").strip().lower()
            events = self.store.list_events(limit=700)
            if q:
                filtered = []
                for ev in events:
                    blob = " ".join(
                        [
                            str(ev.get("timestamp", "")),
                            str(ev.get("action", "")),
                            str(ev.get("entity_type", "")),
                            str(ev.get("entity_id", "")),
                            str(ev.get("details", "")),
                        ]
                    ).lower()
                    if q in blob:
                        filtered.append(ev)
                events = filtered

            for item in self.activity_tree.get_children(""):
                self.activity_tree.delete(item)

            # Newest first
            for ev in reversed(events):
                ts = str(ev.get("timestamp", ""))
                action = str(ev.get("action", ""))
                entity = f"{ev.get('entity_type','')}:{ev.get('entity_id','')}"
                details = str(ev.get("details", ""))
                self.activity_tree.insert("", "end", values=(ts, action, entity, details))
        except Exception as e:
            self.err_logger.log_exception(e, "refresh_activity")

    # ---------------- Selection handlers ----------------
    def _on_student_selected(self, _event=None) -> None:
        sel = self.students_tree.selection()
        if not sel:
            return
        self.select_student(sel[0])

    def _on_teacher_selected(self, _event=None) -> None:
        sel = self.teachers_tree.selection()
        if not sel:
            return
        self.select_teacher(sel[0])

    def select_student(self, student_id: str) -> None:
        self.selected = Selected(entity="student", person_id=student_id)
        self._update_student_payment_label()

    def select_teacher(self, teacher_id: str) -> None:
        self.selected = Selected(entity="teacher", person_id=teacher_id)
        self._update_teacher_payment_label()

    def _update_student_payment_label(self) -> None:
        if self.selected.entity != "student" or not self.selected.person_id:
            self.student_payment_status.configure(text="Status: (select a student)")
            return
        year = _safe_int(self.student_year.get(), self.settings.default_year)
        month = self._month_index(self.student_month.get())
        status = self.store.get_payment_status("student", self.selected.person_id, year, month)
        self.student_payment_status.configure(text=f"Status: {status}  ({MONTHS[month-1]} {year})")

    def _update_teacher_payment_label(self) -> None:
        if self.selected.entity != "teacher" or not self.selected.person_id:
            self.teacher_payment_status.configure(text="Status: (select a teacher)")
            return
        year = _safe_int(self.teacher_year.get(), self.settings.default_year)
        month = self._month_index(self.teacher_month.get())
        status = self.store.get_payment_status("teacher", self.selected.person_id, year, month)
        self.teacher_payment_status.configure(text=f"Status: {status}  ({MONTHS[month-1]} {year})")

    # ---------------- CRUD dialogs ----------------
    def _person_dialog(self, title: str, fields: list[tuple[str, str]], custom_fields: list[str], initial: dict[str, str] | None = None):
        dlg = ctk.CTkToplevel(self)
        dlg.title(title)
        dlg.geometry("620x680")
        dlg.transient(self)
        dlg.grab_set()
        dlg.grid_columnconfigure(0, weight=1)
        dlg.grid_rowconfigure(0, weight=1)

        shell = self._card(dlg)
        shell.grid(row=0, column=0, sticky="nsew", padx=14, pady=14)
        shell.grid_columnconfigure(0, weight=1)
        shell.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(shell, text=title, font=ctk.CTkFont(size=18, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=14, pady=(14, 10)
        )

        body = ctk.CTkScrollableFrame(shell)
        body.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 10))
        body.grid_columnconfigure(1, weight=1)

        entries: dict[str, ctk.CTkEntry] = {}

        def add_row(row: int, key: str, label: str):
            ctk.CTkLabel(body, text=label).grid(row=row, column=0, padx=8, pady=8, sticky="w")
            ent = ctk.CTkEntry(body)
            ent.grid(row=row, column=1, padx=8, pady=8, sticky="ew")
            if initial and key in initial:
                ent.insert(0, str(initial.get(key, "")))
            entries[key] = ent

        r = 0
        for key, label in fields:
            add_row(r, key, label)
            r += 1

        if custom_fields:
            ctk.CTkLabel(body, text="Custom Fields", font=ctk.CTkFont(weight="bold")).grid(
                row=r, column=0, columnspan=2, padx=8, pady=(18, 6), sticky="w"
            )
            r += 1
            for cf in custom_fields:
                add_row(r, cf, cf)
                r += 1

        actions = ctk.CTkFrame(shell, corner_radius=0)
        actions.grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 14))
        actions.grid_columnconfigure((0, 1), weight=1)

        result: dict[str, str] = {}

        def on_save():
            for k, ent in entries.items():
                result[k] = ent.get().strip()
            dlg.destroy()

        def on_cancel():
            dlg.destroy()

        ctk.CTkButton(actions, text="Cancel", command=on_cancel).grid(row=0, column=0, padx=6, pady=10, sticky="ew")
        ctk.CTkButton(actions, text="Save", command=on_save).grid(row=0, column=1, padx=6, pady=10, sticky="ew")

        self.wait_window(dlg)
        return result if result else None

    def _next_id(self, entity: str, prefix: str, existing_ids: list[str]) -> str:
        max_n = 0
        for eid in existing_ids:
            if not eid.startswith(prefix):
                continue
            tail = eid[len(prefix) :]
            m = re.match(r"0*(\d+)$", tail)
            if m:
                max_n = max(max_n, int(m.group(1)))
        return f"{prefix}{max_n + 1:04d}"

    # ---------------- Actions (Students/Teachers) ----------------
    def add_student(self) -> None:
        try:
            students = self.store.list_students()
            existing = [str(s.get("student_id", "")) for s in students]
            sid = self._next_id("student", self.settings.student_id_prefix, existing)

            fields = [
                ("student_id", "Student ID"),
                ("first_name", "First Name"),
                ("last_name", "Last Name"),
                ("class", "Class"),
                ("section", "Section"),
                ("primary_contact", "Primary Contact"),
                ("secondary_contact", "Secondary Contact"),
            ]
            initial = {"student_id": sid}
            res = self._person_dialog("Add Student", fields, self.settings.student_custom_fields, initial=initial)
            if not res:
                return
            self.store.upsert_student(res)
            self._emit(
                "add_student",
                "student",
                res.get("student_id", ""),
                f"{res.get('first_name','')} {res.get('last_name','')}",
            )
            self.refresh_students()
            self.refresh_dashboard()
            self.refresh_activity()
        except Exception as e:
            self.err_logger.log_exception(e, "add_student")

    def edit_selected_student(self) -> None:
        if self.selected.entity != "student" or not self.selected.person_id:
            return
        try:
            students = self.store.list_students()
            current = None
            for s in students:
                if str(s.get("student_id", "")) == self.selected.person_id:
                    current = {k: "" if v is None else str(v) for k, v in s.items()}
                    break
            if not current:
                return

            fields = [
                ("student_id", "Student ID"),
                ("first_name", "First Name"),
                ("last_name", "Last Name"),
                ("class", "Class"),
                ("section", "Section"),
                ("primary_contact", "Primary Contact"),
                ("secondary_contact", "Secondary Contact"),
            ]

            res = self._person_dialog("Edit Student", fields, self.settings.student_custom_fields, initial=current)
            if not res:
                return
            res["student_id"] = self.selected.person_id
            self.store.upsert_student(res)
            self._emit("edit_student", "student", self.selected.person_id, "updated profile")
            self.refresh_students()
            self.refresh_dashboard()
            self.refresh_activity()
        except Exception as e:
            self.err_logger.log_exception(e, "edit_selected_student")

    def delete_selected_student(self) -> None:
        if self.selected.entity != "student" or not self.selected.person_id:
            return
        try:
            ok = self.store.delete_student(self.selected.person_id)
            if ok:
                self._emit("delete_student", "student", self.selected.person_id, "deleted")
            self.selected = Selected()
            self.student_payment_status.configure(text="Status: (select a student)")
            self.refresh_students()
            self.refresh_dashboard()
            self.refresh_activity()
        except Exception as e:
            self.err_logger.log_exception(e, "delete_selected_student")

    def add_teacher(self) -> None:
        try:
            teachers = self.store.list_teachers()
            existing = [str(t.get("teacher_id", "")) for t in teachers]
            tid = self._next_id("teacher", self.settings.teacher_id_prefix, existing)

            fields = [
                ("teacher_id", "Teacher ID"),
                ("first_name", "First Name"),
                ("last_name", "Last Name"),
                ("role", "Role/Position"),
                ("primary_contact", "Primary Contact"),
                ("secondary_contact", "Secondary Contact"),
            ]
            initial = {"teacher_id": tid}
            res = self._person_dialog("Add Teacher", fields, self.settings.teacher_custom_fields, initial=initial)
            if not res:
                return
            self.store.upsert_teacher(res)
            self._emit(
                "add_teacher",
                "teacher",
                res.get("teacher_id", ""),
                f"{res.get('first_name','')} {res.get('last_name','')}",
            )
            self.refresh_teachers()
            self.refresh_dashboard()
            self.refresh_activity()
        except Exception as e:
            self.err_logger.log_exception(e, "add_teacher")

    def edit_selected_teacher(self) -> None:
        if self.selected.entity != "teacher" or not self.selected.person_id:
            return
        try:
            teachers = self.store.list_teachers()
            current = None
            for t in teachers:
                if str(t.get("teacher_id", "")) == self.selected.person_id:
                    current = {k: "" if v is None else str(v) for k, v in t.items()}
                    break
            if not current:
                return

            fields = [
                ("teacher_id", "Teacher ID"),
                ("first_name", "First Name"),
                ("last_name", "Last Name"),
                ("role", "Role/Position"),
                ("primary_contact", "Primary Contact"),
                ("secondary_contact", "Secondary Contact"),
            ]
            res = self._person_dialog("Edit Teacher", fields, self.settings.teacher_custom_fields, initial=current)
            if not res:
                return
            res["teacher_id"] = self.selected.person_id
            self.store.upsert_teacher(res)
            self._emit("edit_teacher", "teacher", self.selected.person_id, "updated profile")
            self.refresh_teachers()
            self.refresh_dashboard()
            self.refresh_activity()
        except Exception as e:
            self.err_logger.log_exception(e, "edit_selected_teacher")

    def delete_selected_teacher(self) -> None:
        if self.selected.entity != "teacher" or not self.selected.person_id:
            return
        try:
            ok = self.store.delete_teacher(self.selected.person_id)
            if ok:
                self._emit("delete_teacher", "teacher", self.selected.person_id, "deleted")
            self.selected = Selected()
            self.teacher_payment_status.configure(text="Status: (select a teacher)")
            self.refresh_teachers()
            self.refresh_dashboard()
            self.refresh_activity()
        except Exception as e:
            self.err_logger.log_exception(e, "delete_selected_teacher")

    # ---------------- Payments ----------------
    def toggle_student_payment(self) -> None:
        if self.selected.entity != "student" or not self.selected.person_id:
            return
        try:
            year = _safe_int(self.student_year.get(), self.settings.default_year)
            month = self._month_index(self.student_month.get())
            current = self.store.get_payment_status("student", self.selected.person_id, year, month)
            new_status = "Pending" if current.lower() == "paid" else "Paid"
            self.store.set_payment_status("student", self.selected.person_id, year, month, new_status)
            self._emit("toggle_tuition", "student", self.selected.person_id, f"{MONTHS[month-1]} {year}: {new_status}")
            self._update_student_payment_label()
            self.refresh_dashboard()
            self.refresh_activity()
        except Exception as e:
            self.err_logger.log_exception(e, "toggle_student_payment")

    def toggle_teacher_payment(self) -> None:
        if self.selected.entity != "teacher" or not self.selected.person_id:
            return
        try:
            year = _safe_int(self.teacher_year.get(), self.settings.default_year)
            month = self._month_index(self.teacher_month.get())
            current = self.store.get_payment_status("teacher", self.selected.person_id, year, month)
            new_status = "Pending" if current.lower() == "paid" else "Paid"
            self.store.set_payment_status("teacher", self.selected.person_id, year, month, new_status)
            self._emit("toggle_salary", "teacher", self.selected.person_id, f"{MONTHS[month-1]} {year}: {new_status}")
            self._update_teacher_payment_label()
            self.refresh_dashboard()
            self.refresh_activity()
        except Exception as e:
            self.err_logger.log_exception(e, "toggle_teacher_payment")

    # ---------------- Dashboard ----------------
    def refresh_dashboard(self) -> None:
        try:
            year = _safe_int(self.dash_year.get(), self.settings.default_year)
            month = self._month_index(self.dash_month.get())

            students = self.store.list_students()
            teachers = self.store.list_teachers()
            self.dash_students.configure(text=f"Students: {len(students)}")
            self.dash_teachers.configure(text=f"Teachers: {len(teachers)}")

            s = self.store.payment_stats("student", year, month)
            t = self.store.payment_stats("teacher", year, month)
            self.dash_students_paid.configure(
                text=f"Tuition (Paid/Pending): {s['paid']} / {s['pending']}  ({MONTHS[month-1]} {year})"
            )
            self.dash_teachers_paid.configure(
                text=f"Salary (Paid/Pending): {t['paid']} / {t['pending']}  ({MONTHS[month-1]} {year})"
            )
        except Exception as e:
            self.err_logger.log_exception(e, "refresh_dashboard")

    # ---------------- Settings ----------------
    def _clear_scrollable(self, frame: ctk.CTkScrollableFrame) -> None:
        for child in frame.winfo_children():
            child.destroy()

    def refresh_custom_fields_views(self) -> None:
        self._clear_scrollable(self.student_cf_list)
        for cf in self.settings.student_custom_fields:
            row = ctk.CTkFrame(self.student_cf_list)
            row.grid(sticky="ew", padx=6, pady=4)
            row.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(row, text=cf, anchor="w").grid(row=0, column=0, padx=8, pady=8, sticky="ew")
            ctk.CTkButton(row, text="Remove", width=90, command=lambda _cf=cf: self.remove_student_custom_field(_cf)).grid(
                row=0, column=1, padx=8, pady=8
            )

        self._clear_scrollable(self.teacher_cf_list)
        for cf in self.settings.teacher_custom_fields:
            row = ctk.CTkFrame(self.teacher_cf_list)
            row.grid(sticky="ew", padx=6, pady=4)
            row.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(row, text=cf, anchor="w").grid(row=0, column=0, padx=8, pady=8, sticky="ew")
            ctk.CTkButton(row, text="Remove", width=90, command=lambda _cf=cf: self.remove_teacher_custom_field(_cf)).grid(
                row=0, column=1, padx=8, pady=8
            )

    def add_student_custom_field(self) -> None:
        raw = self.student_cf_entry.get()
        name = _normalize_field_name(raw)
        if not name:
            return
        if name in self.settings.student_custom_fields:
            return
        self.settings.student_custom_fields.append(name)
        self.student_cf_entry.delete(0, "end")
        self.save_settings(rebuild_workbook=True)

    def add_teacher_custom_field(self) -> None:
        raw = self.teacher_cf_entry.get()
        name = _normalize_field_name(raw)
        if not name:
            return
        if name in self.settings.teacher_custom_fields:
            return
        self.settings.teacher_custom_fields.append(name)
        self.teacher_cf_entry.delete(0, "end")
        self.save_settings(rebuild_workbook=True)

    def remove_student_custom_field(self, name: str) -> None:
        if name in self.settings.student_custom_fields:
            self.settings.student_custom_fields.remove(name)
            self.save_settings(rebuild_workbook=True)

    def remove_teacher_custom_field(self, name: str) -> None:
        if name in self.settings.teacher_custom_fields:
            self.settings.teacher_custom_fields.remove(name)
            self.save_settings(rebuild_workbook=True)

    def save_settings(self, rebuild_workbook: bool = False) -> None:
        try:
            self.settings.student_id_prefix = (self.student_prefix.get() or "STU-").strip() or "STU-"
            self.settings.teacher_id_prefix = (self.teacher_prefix.get() or "TCH-").strip() or "TCH-"
            self.settings.appearance_mode = (self.appearance_mode.get() or "System").strip().capitalize()
            try:
                self.settings.ui_scaling = float(self.ui_scale.get() or 1.0)
            except Exception:
                self.settings.ui_scaling = 1.0

            self.settings_store.save(self.settings)
            if rebuild_workbook:
                self.store.ensure_workbook(self.settings.student_custom_fields, self.settings.teacher_custom_fields)

            self._apply_ui_settings()
            self._emit("save_settings", "settings", "local", "updated settings.json")
            self.refresh_custom_fields_views()
            self.refresh_activity()
        except Exception as e:
            self.err_logger.log_exception(e, "save_settings")

    # ---------------- Close ----------------
    def _on_close(self) -> None:
        # Nothing special for now, but this is a good place for future cleanup.
        self.destroy()


def run_app() -> None:
    _enable_high_dpi()
    app = HTSMSApp()
    app.mainloop()
