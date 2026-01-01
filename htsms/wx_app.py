from __future__ import annotations

import sys
from dataclasses import dataclass
from typing import Any

import wx
import wx.adv
import wx.dataview as dv

from .constants import APP_NAME
from .logger import ErrorLogger, now_ts, AppEvent
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


def _safe_int(value: str, default: int = 0) -> int:
    try:
        return int(str(value).strip())
    except Exception:
        return default


@dataclass
class Selected:
    entity: str = ""
    person_id: str = ""


class CardPanel(wx.Panel):
    """A simple modern-looking card panel (rounded border + padding)."""

    def __init__(self, parent: wx.Window, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.SetBackgroundStyle(wx.BG_STYLE_PAINT)
        self.Bind(wx.EVT_PAINT, self._on_paint)

    def _on_paint(self, evt: wx.PaintEvent):
        dc = wx.AutoBufferedPaintDC(self)
        dc.Clear()
        gc = wx.GraphicsContext.Create(dc)
        if not gc:
            return

        rect = self.GetClientRect()
        pad = 2
        r = wx.Rect(rect.x + pad, rect.y + pad, rect.width - 2 * pad, rect.height - 2 * pad)

        is_dark = wx.SystemSettings.GetAppearance().IsDark()
        bg = wx.Colour(17, 24, 39) if is_dark else wx.Colour(255, 255, 255)
        border = wx.Colour(31, 41, 55) if is_dark else wx.Colour(229, 231, 235)

        gc.SetPen(wx.Pen(border, 1))
        gc.SetBrush(wx.Brush(bg))
        gc.DrawRoundedRectangle(r.x, r.y, r.width, r.height, 12)


class SidebarButton(wx.Control):
    """Custom sidebar button (fast, crisp, commercial-like)."""

    def __init__(self, parent: wx.Window, label: str, key: str):
        super().__init__(parent, style=wx.BORDER_NONE)
        self.key = key
        self.label = label
        self.active = False

        self.SetMinSize((180, 44))
        self.SetBackgroundStyle(wx.BG_STYLE_PAINT)
        self.Bind(wx.EVT_PAINT, self._on_paint)
        self.Bind(wx.EVT_LEFT_UP, self._on_click)
        self.Bind(wx.EVT_ENTER_WINDOW, lambda e: (self.SetCursor(wx.Cursor(wx.CURSOR_HAND)), e.Skip()))

    def _on_click(self, _evt: wx.MouseEvent):
        evt = wx.CommandEvent(wx.EVT_BUTTON.typeId, self.GetId())
        evt.SetString(self.key)
        wx.PostEvent(self.GetParent(), evt)

    def _on_paint(self, evt: wx.PaintEvent):
        dc = wx.AutoBufferedPaintDC(self)
        dc.Clear()
        gc = wx.GraphicsContext.Create(dc)
        if not gc:
            return

        rect = self.GetClientRect()
        is_dark = wx.SystemSettings.GetAppearance().IsDark()

        base = wx.Colour(15, 23, 42) if is_dark else wx.Colour(248, 250, 252)
        gc.SetBrush(wx.Brush(base))
        gc.SetPen(wx.TRANSPARENT_PEN)
        gc.DrawRectangle(rect.x, rect.y, rect.width, rect.height)

        if self.active:
            accent = wx.Colour(59, 130, 246)  # blue
            fill = wx.Colour(11, 59, 112) if is_dark else wx.Colour(219, 234, 254)
            gc.SetBrush(wx.Brush(fill))
            gc.SetPen(wx.Pen(accent, 1))
            gc.DrawRoundedRectangle(rect.x + 8, rect.y + 6, rect.width - 16, rect.height - 12, 10)
            text_col = wx.Colour(255, 255, 255) if is_dark else wx.Colour(17, 24, 39)
        else:
            text_col = wx.Colour(255, 255, 255) if is_dark else wx.Colour(17, 24, 39)

        gc.SetFont(wx.SystemSettings.GetFont(wx.SYS_DEFAULT_GUI_FONT).Bold(), text_col)
        gc.DrawText(self.label, rect.x + 18, rect.y + 13)


class StudentDialog(wx.Dialog):
    def __init__(
        self,
        parent: wx.Window,
        *,
        title: str,
        student_id: str,
        initial: dict[str, Any] | None = None,
        custom_fields: list[str] | None = None,
    ):
        super().__init__(parent, title=title, style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.SetMinSize((560, 520))

        self.student_id = student_id
        self.initial = initial or {}
        self.custom_fields = custom_fields or []

        root = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)
        root.SetSizer(sizer)

        header = wx.StaticText(root, label=title)
        header.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        sizer.Add(header, 0, wx.LEFT | wx.RIGHT | wx.TOP, 16)
        sizer.Add(wx.StaticText(root, label=f"Student ID: {student_id}"), 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 16)

        form = wx.FlexGridSizer(rows=0, cols=2, vgap=10, hgap=12)
        form.AddGrowableCol(1, 1)

        def add_row(lbl: str, ctrl: wx.Window):
            form.Add(wx.StaticText(root, label=lbl), 0, wx.ALIGN_CENTER_VERTICAL)
            form.Add(ctrl, 1, wx.EXPAND)

        self.txt_first = wx.TextCtrl(root, value=str(self.initial.get("first_name", "") or ""))
        self.txt_last = wx.TextCtrl(root, value=str(self.initial.get("last_name", "") or ""))

        # Age: store as empty when unknown.
        age_val = str(self.initial.get("age", "") or "").strip()
        age_int = _safe_int(age_val, default=0)
        self.spin_age = wx.SpinCtrl(root, min=0, max=120, initial=age_int)
        self.spin_age.SetToolTip("Set to 0 if unknown")

        self.txt_class = wx.TextCtrl(root, value=str(self.initial.get("class", "") or ""))
        self.txt_section = wx.TextCtrl(root, value=str(self.initial.get("section", "") or ""))
        self.txt_p1 = wx.TextCtrl(root, value=str(self.initial.get("primary_contact", "") or ""))
        self.txt_p2 = wx.TextCtrl(root, value=str(self.initial.get("secondary_contact", "") or ""))

        add_row("First name *", self.txt_first)
        add_row("Last name", self.txt_last)
        add_row("Age", self.spin_age)
        add_row("Class", self.txt_class)
        add_row("Section", self.txt_section)
        add_row("Primary contact", self.txt_p1)
        add_row("Secondary contact", self.txt_p2)

        self._custom_ctrls: dict[str, wx.TextCtrl] = {}
        if self.custom_fields:
            sizer.Add(wx.StaticLine(root), 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, 16)
            sizer.Add(wx.StaticText(root, label="Custom fields"), 0, wx.LEFT | wx.RIGHT | wx.TOP, 16)
            for field_name in self.custom_fields:
                ctrl = wx.TextCtrl(root, value=str(self.initial.get(field_name, "") or ""))
                self._custom_ctrls[field_name] = ctrl
                add_row(field_name, ctrl)

        sizer.Add(form, 1, wx.EXPAND | wx.LEFT | wx.RIGHT, 16)

        btns = self.CreateSeparatedButtonSizer(wx.OK | wx.CANCEL)
        if btns:
            sizer.Add(btns, 0, wx.EXPAND | wx.ALL, 12)

        self.Bind(wx.EVT_BUTTON, self._on_ok, id=wx.ID_OK)
        self.Fit()
        self.Layout()
        self.CentreOnParent()

    def _on_ok(self, evt: wx.CommandEvent) -> None:
        first = self.txt_first.GetValue().strip()
        if not first:
            wx.MessageBox("First name is required.", "Validation", wx.OK | wx.ICON_WARNING)
            self.txt_first.SetFocus()
            return
        evt.Skip()

    def get_data(self) -> dict[str, Any]:
        age = int(self.spin_age.GetValue())
        data: dict[str, Any] = {
            "student_id": self.student_id,
            "first_name": self.txt_first.GetValue().strip(),
            "last_name": self.txt_last.GetValue().strip(),
            "age": "" if age <= 0 else age,
            "class": self.txt_class.GetValue().strip(),
            "section": self.txt_section.GetValue().strip(),
            "primary_contact": self.txt_p1.GetValue().strip(),
            "secondary_contact": self.txt_p2.GetValue().strip(),
        }
        for k, c in self._custom_ctrls.items():
            data[k] = c.GetValue().strip()
        return data


class HTSMSFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title=APP_NAME, size=(1280, 800))

        self.err_logger = ErrorLogger()
        self.settings_store = SettingsStore()
        self.settings = self.settings_store.load()

        self.store = ExcelStore()
        self.store.ensure_workbook(self.settings.student_custom_fields, self.settings.teacher_custom_fields)

        self.selected = Selected()

        self._build_ui()
        self._wire_events()
        self.refresh_all()
        self.show_page("dashboard")

    def _build_ui(self) -> None:
        self.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))

        root = wx.Panel(self)
        root_sizer = wx.BoxSizer(wx.HORIZONTAL)
        root.SetSizer(root_sizer)

        # Sidebar
        self.sidebar = wx.Panel(root)
        self.sidebar.SetMinSize((220, -1))
        self.sidebar.SetBackgroundColour(wx.Colour(248, 250, 252))
        sb = wx.BoxSizer(wx.VERTICAL)
        self.sidebar.SetSizer(sb)

        title = wx.StaticText(self.sidebar, label="High Tech\nSchool Management")
        title.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        sb.Add(title, 0, wx.ALL, 16)

        self.btn_dashboard = SidebarButton(self.sidebar, "Dashboard", "dashboard")
        self.btn_students = SidebarButton(self.sidebar, "Students", "students")
        self.btn_teachers = SidebarButton(self.sidebar, "Teachers", "teachers")
        self.btn_activity = SidebarButton(self.sidebar, "Activity Log", "activity")
        self.btn_settings = SidebarButton(self.sidebar, "Settings", "settings")

        for b in [self.btn_dashboard, self.btn_students, self.btn_teachers, self.btn_activity, self.btn_settings]:
            sb.Add(b, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 8)
            sb.AddSpacer(6)

        sb.AddStretchSpacer(1)
        sb.Add(wx.StaticText(self.sidebar, label="Data: school_data.xlsx\nLogs: error_log.txt"), 0, wx.ALL, 16)

        # Content area
        self.content = wx.Panel(root)
        content_sizer = wx.BoxSizer(wx.VERTICAL)
        self.content.SetSizer(content_sizer)

        self.page_title = wx.StaticText(self.content, label="")
        self.page_title.SetFont(wx.Font(18, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        content_sizer.Add(self.page_title, 0, wx.LEFT | wx.TOP, 18)

        self.page_subtitle = wx.StaticText(self.content, label="")
        content_sizer.Add(self.page_subtitle, 0, wx.LEFT | wx.BOTTOM, 18)

        self.pages = wx.Simplebook(self.content)
        content_sizer.Add(self.pages, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 18)

        # Dashboard page (placeholder card layout; charts added next)
        self.pg_dashboard = wx.Panel(self.pages)
        dash = wx.BoxSizer(wx.VERTICAL)
        self.pg_dashboard.SetSizer(dash)

        row = wx.BoxSizer(wx.HORIZONTAL)
        self.dash_card_students = CardPanel(self.pg_dashboard)
        self.dash_card_teachers = CardPanel(self.pg_dashboard)
        row.Add(self.dash_card_students, 1, wx.EXPAND | wx.RIGHT, 10)
        row.Add(self.dash_card_teachers, 1, wx.EXPAND | wx.LEFT, 10)
        dash.Add(row, 0, wx.EXPAND | wx.BOTTOM, 12)

        self.lbl_students = wx.StaticText(self.dash_card_students, label="Students: 0")
        self.lbl_students.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        s1 = wx.BoxSizer(wx.VERTICAL)
        s1.Add(self.lbl_students, 0, wx.ALL, 16)
        self.dash_card_students.SetSizer(s1)

        self.lbl_teachers = wx.StaticText(self.dash_card_teachers, label="Teachers: 0")
        self.lbl_teachers.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        s2 = wx.BoxSizer(wx.VERTICAL)
        s2.Add(self.lbl_teachers, 0, wx.ALL, 16)
        self.dash_card_teachers.SetSizer(s2)

        # Students page
        self.pg_students = wx.Panel(self.pages)
        sp = wx.BoxSizer(wx.VERTICAL)
        self.pg_students.SetSizer(sp)

        tools = CardPanel(self.pg_students)
        tools_s = wx.BoxSizer(wx.HORIZONTAL)
        tools.SetSizer(tools_s)

        self.student_search = wx.SearchCtrl(tools, style=wx.TE_PROCESS_ENTER)
        self.student_search.SetDescriptiveText("Search students (name, ID, class, section, age, contact)")

        self.student_filter_field = wx.Choice(tools, choices=["All", "Name", "ID", "Class", "Section", "Age", "Contact"])
        self.student_filter_field.SetSelection(0)

        self.student_filter_class = wx.Choice(tools, choices=["(All classes)"])
        self.student_filter_class.SetSelection(0)

        self.student_filter_section = wx.Choice(tools, choices=["(All sections)"])
        self.student_filter_section.SetSelection(0)

        self.btn_add_student = wx.Button(tools, label="Add Student")
        self.btn_edit_student = wx.Button(tools, label="Edit Selected")
        self.btn_del_student = wx.Button(tools, label="Delete")

        tools_s.Add(self.student_search, 2, wx.ALL | wx.EXPAND, 12)
        tools_s.Add(self.student_filter_field, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 12)
        tools_s.Add(self.student_filter_class, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 12)
        tools_s.Add(self.student_filter_section, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 12)
        tools_s.AddStretchSpacer(1)
        tools_s.Add(self.btn_add_student, 0, wx.ALL, 12)
        tools_s.Add(self.btn_edit_student, 0, wx.ALL, 12)
        tools_s.Add(self.btn_del_student, 0, wx.ALL, 12)

        sp.Add(tools, 0, wx.EXPAND | wx.BOTTOM, 12)

        table_card = CardPanel(self.pg_students)
        tc = wx.BoxSizer(wx.VERTICAL)
        table_card.SetSizer(tc)

        self.students_dv = dv.DataViewListCtrl(table_card, style=dv.DV_ROW_LINES | dv.DV_VERT_RULES)
        self.students_dv.AppendTextColumn("ID", width=110)
        self.students_dv.AppendTextColumn("Name", width=220)
        self.students_dv.AppendTextColumn("Age", width=60)
        self.students_dv.AppendTextColumn("Class", width=70)
        self.students_dv.AppendTextColumn("Section", width=80)
        self.students_dv.AppendTextColumn("Primary", width=160)
        self.students_dv.AppendTextColumn("Secondary", width=160)

        tc.Add(self.students_dv, 1, wx.EXPAND | wx.ALL, 12)
        sp.Add(table_card, 1, wx.EXPAND)

        # Teachers page (placeholder table)
        self.pg_teachers = wx.Panel(self.pages)
        tp = wx.BoxSizer(wx.VERTICAL)
        self.pg_teachers.SetSizer(tp)
        t_card = CardPanel(self.pg_teachers)
        t_s = wx.BoxSizer(wx.VERTICAL)
        t_card.SetSizer(t_s)
        self.teachers_dv = dv.DataViewListCtrl(t_card, style=dv.DV_ROW_LINES | dv.DV_VERT_RULES)
        self.teachers_dv.AppendTextColumn("ID", width=110)
        self.teachers_dv.AppendTextColumn("Name", width=220)
        self.teachers_dv.AppendTextColumn("Role", width=180)
        self.teachers_dv.AppendTextColumn("Primary", width=160)
        self.teachers_dv.AppendTextColumn("Secondary", width=160)
        t_s.Add(self.teachers_dv, 1, wx.EXPAND | wx.ALL, 12)
        tp.Add(t_card, 1, wx.EXPAND)

        # Activity page (placeholder)
        self.pg_activity = wx.Panel(self.pages)
        ap = wx.BoxSizer(wx.VERTICAL)
        self.pg_activity.SetSizer(ap)
        a_card = CardPanel(self.pg_activity)
        a_s = wx.BoxSizer(wx.VERTICAL)
        a_card.SetSizer(a_s)
        self.activity_dv = dv.DataViewListCtrl(a_card, style=dv.DV_ROW_LINES | dv.DV_VERT_RULES)
        self.activity_dv.AppendTextColumn("Timestamp", width=160)
        self.activity_dv.AppendTextColumn("Action", width=140)
        self.activity_dv.AppendTextColumn("Entity", width=160)
        self.activity_dv.AppendTextColumn("Details", width=520)
        a_s.Add(self.activity_dv, 1, wx.EXPAND | wx.ALL, 12)
        ap.Add(a_card, 1, wx.EXPAND)

        # Settings page (placeholder)
        self.pg_settings = wx.Panel(self.pages)
        st = wx.BoxSizer(wx.VERTICAL)
        self.pg_settings.SetSizer(st)
        st.Add(wx.StaticText(self.pg_settings, label="Settings UI is next (prefixes, custom fields, appearance, scaling)."), 0, wx.ALL, 16)

        self.pages.AddPage(self.pg_dashboard, "dashboard")
        self.pages.AddPage(self.pg_students, "students")
        self.pages.AddPage(self.pg_teachers, "teachers")
        self.pages.AddPage(self.pg_activity, "activity")
        self.pages.AddPage(self.pg_settings, "settings")

        root_sizer.Add(self.sidebar, 0, wx.EXPAND)
        root_sizer.Add(self.content, 1, wx.EXPAND)

        self.CreateStatusBar()
        self.SetStatusText("Ready")

    def _wire_events(self) -> None:
        # Sidebar clicks
        self.sidebar.Bind(wx.EVT_BUTTON, self._on_nav)

        # Students actions
        self.btn_add_student.Bind(wx.EVT_BUTTON, lambda e: self.add_student())
        self.btn_edit_student.Bind(wx.EVT_BUTTON, lambda e: self.edit_selected_student())
        self.btn_del_student.Bind(wx.EVT_BUTTON, lambda e: self.delete_selected_student())

        self.student_search.Bind(wx.EVT_TEXT, lambda e: self.refresh_students())
        self.student_filter_field.Bind(wx.EVT_CHOICE, lambda e: self.refresh_students())
        self.student_filter_class.Bind(wx.EVT_CHOICE, lambda e: self.refresh_students())
        self.student_filter_section.Bind(wx.EVT_CHOICE, lambda e: self.refresh_students())

        self.students_dv.Bind(dv.EVT_DATAVIEW_SELECTION_CHANGED, self._on_student_selected)
        self.teachers_dv.Bind(dv.EVT_DATAVIEW_SELECTION_CHANGED, self._on_teacher_selected)

    def _emit(self, action: str, entity_type: str, entity_id: str, details: str = "") -> None:
        self.store.add_event(AppEvent(timestamp=now_ts(), action=action, entity_type=entity_type, entity_id=entity_id, details=details))

    def _on_nav(self, evt: wx.CommandEvent) -> None:
        key = evt.GetString()
        if key:
            self.show_page(key)

    def show_page(self, key: str) -> None:
        # Button active state
        for b in [self.btn_dashboard, self.btn_students, self.btn_teachers, self.btn_activity, self.btn_settings]:
            b.active = (b.key == key)
            b.Refresh(False)

        # Page title
        titles = {
            "dashboard": ("Dashboard", "Overview of students, teachers, and payments"),
            "students": ("Students", "Manage profiles, class/section grouping, and age"),
            "teachers": ("Teachers", "Manage teacher profiles"),
            "activity": ("Activity Log", "Every action is recorded"),
            "settings": ("Settings", "Prefixes, fields, and UI options"),
        }
        t, s = titles.get(key, ("", ""))
        self.page_title.SetLabel(t)
        self.page_subtitle.SetLabel(s)

        idx = {"dashboard": 0, "students": 1, "teachers": 2, "activity": 3, "settings": 4}.get(key, 0)
        self.pages.SetSelection(idx)

        # Lightweight refresh
        if key == "dashboard":
            self.refresh_dashboard()
        elif key == "students":
            self.refresh_students()
        elif key == "teachers":
            self.refresh_teachers()
        elif key == "activity":
            self.refresh_activity()

    # --------- Refresh methods (realtime Excel -> UI) ---------
    def refresh_all(self) -> None:
        self.refresh_dashboard()
        self.refresh_students_filters()
        self.refresh_students()
        self.refresh_teachers()
        self.refresh_activity()

    def refresh_dashboard(self) -> None:
        try:
            students = self.store.list_students()
            teachers = self.store.list_teachers()
            self.lbl_students.SetLabel(f"Students: {len(students)}")
            self.lbl_teachers.SetLabel(f"Teachers: {len(teachers)}")
        except Exception as e:
            self.err_logger.log_exception(e, "wx_refresh_dashboard")

    def refresh_students_filters(self) -> None:
        try:
            rows = self.store.list_students()
            classes = sorted({str(r.get("class", "")).strip() for r in rows if str(r.get("class", "")).strip()})
            sections = sorted({str(r.get("section", "")).strip() for r in rows if str(r.get("section", "")).strip()})

            def refill(choice: wx.Choice, first: str, values: list[str]):
                cur = choice.GetStringSelection()
                choice.Clear()
                choice.Append(first)
                for v in values:
                    choice.Append(v)
                if cur and cur in [first] + values:
                    choice.SetStringSelection(cur)
                else:
                    choice.SetSelection(0)

            refill(self.student_filter_class, "(All classes)", classes)
            refill(self.student_filter_section, "(All sections)", sections)
        except Exception as e:
            self.err_logger.log_exception(e, "wx_refresh_students_filters")

    def refresh_students(self) -> None:
        try:
            q = (self.student_search.GetValue() or "").strip().lower()
            field = self.student_filter_field.GetStringSelection() or "All"
            cls_filter = self.student_filter_class.GetStringSelection()
            sec_filter = self.student_filter_section.GetStringSelection()

            rows = self.store.list_students()

            def match_row(r: dict[str, Any]) -> bool:
                if cls_filter and cls_filter != "(All classes)":
                    if str(r.get("class", "")).strip() != cls_filter:
                        return False
                if sec_filter and sec_filter != "(All sections)":
                    if str(r.get("section", "")).strip() != sec_filter:
                        return False

                if not q:
                    return True

                blob_map = {
                    "Name": f"{r.get('first_name','')} {r.get('last_name','')}",
                    "ID": str(r.get("student_id", "")),
                    "Class": str(r.get("class", "")),
                    "Section": str(r.get("section", "")),
                    "Age": str(r.get("age", "")),
                    "Contact": f"{r.get('primary_contact','')} {r.get('secondary_contact','')}",
                    "All": " ".join(
                        [
                            str(r.get("student_id", "")),
                            str(r.get("first_name", "")),
                            str(r.get("last_name", "")),
                            str(r.get("age", "")),
                            str(r.get("class", "")),
                            str(r.get("section", "")),
                            str(r.get("primary_contact", "")),
                            str(r.get("secondary_contact", "")),
                        ]
                    ),
                }
                blob = str(blob_map.get(field, blob_map["All"]))
                return q in blob.lower()

            filtered = [r for r in rows if match_row(r)]
            filtered.sort(key=lambda r: str(r.get("student_id", "")))

            self.students_dv.DeleteAllItems()
            for r in filtered:
                sid = str(r.get("student_id", "")).strip()
                if not sid:
                    continue
                name = f"{r.get('first_name','')} {r.get('last_name','')}".strip()
                age = str(r.get("age", "") or "")
                cls = str(r.get("class", "") or "")
                sec = str(r.get("section", "") or "")
                p1 = str(r.get("primary_contact", "") or "")
                p2 = str(r.get("secondary_contact", "") or "")
                self.students_dv.AppendItem([sid, name, age, cls, sec, p1, p2])

            # Keep selection stable if possible
            if self.selected.entity == "student" and self.selected.person_id:
                self._select_student_in_view(self.selected.person_id)

            self.SetStatusText(f"Students: {len(filtered)} shown")
        except Exception as e:
            self.err_logger.log_exception(e, "wx_refresh_students")

    def _select_student_in_view(self, student_id: str) -> None:
        # dv.DataViewListCtrl doesn't support iid; select by scanning rows.
        for i in range(self.students_dv.GetItemCount()):
            if self.students_dv.GetTextValue(i, 0) == student_id:
                item = self.students_dv.RowToItem(i)
                self.students_dv.Select(item)
                break

    def refresh_teachers(self) -> None:
        try:
            rows = self.store.list_teachers()
            rows.sort(key=lambda r: str(r.get("teacher_id", "")))
            self.teachers_dv.DeleteAllItems()
            for r in rows:
                tid = str(r.get("teacher_id", "")).strip()
                if not tid:
                    continue
                name = f"{r.get('first_name','')} {r.get('last_name','')}".strip()
                role = str(r.get("role", "") or "")
                p1 = str(r.get("primary_contact", "") or "")
                p2 = str(r.get("secondary_contact", "") or "")
                self.teachers_dv.AppendItem([tid, name, role, p1, p2])
        except Exception as e:
            self.err_logger.log_exception(e, "wx_refresh_teachers")

    def refresh_activity(self) -> None:
        try:
            rows = self.store.list_events(limit=500)
            self.activity_dv.DeleteAllItems()
            for ev in reversed(rows):
                ts = str(ev.get("timestamp", ""))
                action = str(ev.get("action", ""))
                entity = f"{ev.get('entity_type','')}:{ev.get('entity_id','')}"
                details = str(ev.get("details", ""))
                self.activity_dv.AppendItem([ts, action, entity, details])
        except Exception as e:
            self.err_logger.log_exception(e, "wx_refresh_activity")

    # --------- Selection handlers ---------
    def _on_student_selected(self, evt: dv.DataViewEvent) -> None:
        try:
            item = evt.GetItem()
            if not item.IsOk():
                return
            row = self.students_dv.ItemToRow(item)
            sid = self.students_dv.GetTextValue(row, 0)
            if sid:
                self.selected = Selected(entity="student", person_id=sid)
        except Exception as e:
            self.err_logger.log_exception(e, "wx_on_student_selected")

    def _on_teacher_selected(self, evt: dv.DataViewEvent) -> None:
        try:
            item = evt.GetItem()
            if not item.IsOk():
                return
            row = self.teachers_dv.ItemToRow(item)
            tid = self.teachers_dv.GetTextValue(row, 0)
            if tid:
                self.selected = Selected(entity="teacher", person_id=tid)
        except Exception as e:
            self.err_logger.log_exception(e, "wx_on_teacher_selected")

    # --------- CRUD dialogs (basic; upgraded next) ---------
    def _next_id(self, entity: str, prefix: str, existing_ids: list[str]) -> str:
        max_n = 0
        for eid in existing_ids:
            if not eid.startswith(prefix):
                continue
            tail = eid[len(prefix) :]
            digits = "".join(ch for ch in tail if ch.isdigit())
            if digits:
                try:
                    max_n = max(max_n, int(digits))
                except Exception:
                    pass
        return f"{prefix}{max_n + 1:04d}"

    def _get_student_by_id(self, student_id: str) -> dict[str, Any] | None:
        for r in self.store.list_students():
            if str(r.get("student_id", "")).strip() == student_id:
                return r
        return None

    def add_student(self) -> None:
        try:
            existing = [str(s.get("student_id", "")) for s in self.store.list_students()]
            sid = self._next_id("student", self.settings.student_id_prefix, existing)

            dlg = StudentDialog(
                self,
                title="Add Student",
                student_id=sid,
                initial=None,
                custom_fields=self.settings.student_custom_fields,
            )
            try:
                if dlg.ShowModal() != wx.ID_OK:
                    return
                data = dlg.get_data()
            finally:
                dlg.Destroy()

            self.store.upsert_student(data)
            self._emit("add_student", "student", sid, f"{data.get('first_name','')} {data.get('last_name','')}".strip())

            # realtime refresh
            self.refresh_students_filters()
            self.refresh_students()
            self.refresh_dashboard()
            self.refresh_activity()

            self.selected = Selected(entity="student", person_id=sid)
            self._select_student_in_view(sid)
        except Exception as e:
            self.err_logger.log_exception(e, "wx_add_student")

    def edit_selected_student(self) -> None:
        if self.selected.entity != "student" or not self.selected.person_id:
            return
        try:
            sid = self.selected.person_id
            current = self._get_student_by_id(sid) or {"student_id": sid}
            dlg = StudentDialog(
                self,
                title="Edit Student",
                student_id=sid,
                initial=current,
                custom_fields=self.settings.student_custom_fields,
            )
            try:
                if dlg.ShowModal() != wx.ID_OK:
                    return
                data = dlg.get_data()
            finally:
                dlg.Destroy()

            self.store.upsert_student(data)
            self._emit("edit_student", "student", sid, "updated")

            self.refresh_students_filters()
            self.refresh_students()
            self.refresh_dashboard()
            self.refresh_activity()
            self._select_student_in_view(sid)
        except Exception as e:
            self.err_logger.log_exception(e, "wx_edit_student")

    def delete_selected_student(self) -> None:
        if self.selected.entity != "student" or not self.selected.person_id:
            return
        try:
            sid = self.selected.person_id
            if wx.MessageBox(f"Delete student {sid}?", "Confirm", wx.YES_NO | wx.ICON_WARNING) != wx.YES:
                return
            ok = self.store.delete_student(sid)
            if ok:
                self._emit("delete_student", "student", sid, "deleted")

            self.selected = Selected()
            self.refresh_students_filters()
            self.refresh_students()
            self.refresh_dashboard()
            self.refresh_activity()
        except Exception as e:
            self.err_logger.log_exception(e, "wx_delete_student")


class HTSMSWxApp(wx.App):
    def __init__(self):
        super().__init__(redirect=False)
        self.err_logger = ErrorLogger()

    def OnInit(self) -> bool:
        try:
            self.frame = HTSMSFrame()
            self.frame.Show()
            return True
        except Exception as e:
            self.err_logger.log_exception(e, "wx_OnInit")
            raise


def run_wx_app() -> None:
    # wx handles DPI pretty well on Windows; keep it simple.
    app = HTSMSWxApp()
    app.MainLoop()
