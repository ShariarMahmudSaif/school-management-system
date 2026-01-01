from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from PySide6 import QtCore, QtGui, QtWidgets
from PySide6.QtCharts import QBarCategoryAxis, QBarSeries, QBarSet, QChart, QChartView, QPieSeries, QValueAxis

from .constants import APP_NAME, DATA_XLSX_PATH
from .logger import AppEvent, ErrorLogger, now_ts
from .settings_store import Settings, SettingsStore
from .storage import ExcelStore
from .qt_pages import PaymentDialog, PaymentsPage, ActivityPage, SettingsPage, MONTHS as MONTHS_FROM_PAGES

MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]


def _app_dark_palette() -> QtGui.QPalette:
    p = QtGui.QPalette()
    # Base theme: dark slate (inspired by modern dashboards)
    p.setColor(QtGui.QPalette.Window, QtGui.QColor(12, 14, 18))
    p.setColor(QtGui.QPalette.WindowText, QtGui.QColor(236, 240, 244))
    p.setColor(QtGui.QPalette.Base, QtGui.QColor(18, 21, 27))
    p.setColor(QtGui.QPalette.AlternateBase, QtGui.QColor(14, 16, 21))
    p.setColor(QtGui.QPalette.ToolTipBase, QtGui.QColor(40, 44, 52))
    p.setColor(QtGui.QPalette.ToolTipText, QtGui.QColor(236, 240, 244))
    p.setColor(QtGui.QPalette.Text, QtGui.QColor(236, 240, 244))
    p.setColor(QtGui.QPalette.Button, QtGui.QColor(22, 25, 32))
    p.setColor(QtGui.QPalette.ButtonText, QtGui.QColor(236, 240, 244))
    p.setColor(QtGui.QPalette.BrightText, QtGui.QColor(255, 90, 90))
    p.setColor(QtGui.QPalette.Highlight, QtGui.QColor(255, 177, 0))  # amber accent
    p.setColor(QtGui.QPalette.HighlightedText, QtGui.QColor(12, 14, 18))
    return p


def _qss() -> str:
    # Keep it crisp: no blur effects; only clean borders, spacing and contrast.
    return """
    QWidget { font-size: 12px; }

    QFrame#Sidebar { background: #0b0d11; border-right: 1px solid #222733; }
    QLabel#AppTitle { font-size: 16px; font-weight: 700; color: #e7edf4; }
    QLabel#SectionTitle { font-size: 12px; font-weight: 700; color: #cfd6df; }

    QPushButton.NavBtn {
        text-align: left;
        padding: 10px 12px;
        border-radius: 10px;
        border: 1px solid transparent;
        color: #e7edf4;
        background: transparent;
    }
    QPushButton.NavBtn:hover { background: #131722; }
    QPushButton.NavBtn[active="true"] {
        background: #1a1f2b;
        border: 1px solid #2a3243;
    }

    QFrame#TopBar { background: #0b0d11; border-bottom: 1px solid #222733; }
    QLabel#TopTitle { font-size: 14px; font-weight: 700; color: #e7edf4; }

    QFrame.Card {
        background: #0f1218;
        border: 1px solid #222733;
        border-radius: 14px;
    }
    QLabel.CardValue { font-size: 18px; font-weight: 800; }
    QLabel.CardLabel { color: #aab3c2; }

    QLineEdit, QComboBox {
        background: #10131a;
        border: 1px solid #222733;
        border-radius: 10px;
        padding: 8px 10px;
    }
    QComboBox::drop-down { border: 0px; }

    QPushButton.Primary {
        background: #ffb100;
        color: #0b0d11;
        border: 0px;
        padding: 10px 12px;
        border-radius: 10px;
        font-weight: 700;
    }
    QPushButton.Primary:hover { background: #ffc14d; }

    QPushButton.Danger {
        background: #e11d48;
        color: white;
        border: 0px;
        padding: 10px 12px;
        border-radius: 10px;
        font-weight: 700;
    }
    QPushButton.Danger:hover { background: #fb2c61; }

    QTableView {
        background: #0f1218;
        border: 1px solid #222733;
        border-radius: 14px;
        gridline-color: #1f2431;
        selection-background-color: #ffb100;
        selection-color: #0b0d11;
    }
    QHeaderView::section {
        background: #0b0d11;
        color: #cfd6df;
        border: 0px;
        padding: 8px;
        font-weight: 700;
    }
    """


@dataclass
class Selected:
    entity: str = ""
    person_id: str = ""


class StudentTableModel(QtCore.QAbstractTableModel):
    COLUMNS = ["ID", "Name", "Age", "Class", "Section", "Primary", "Secondary"]

    def __init__(self, rows: list[dict[str, Any]] | None = None):
        super().__init__()
        self._rows: list[dict[str, Any]] = rows or []

    def set_rows(self, rows: list[dict[str, Any]]) -> None:
        self.beginResetModel()
        self._rows = rows
        self.endResetModel()

    def rowCount(self, parent: QtCore.QModelIndex = QtCore.QModelIndex()) -> int:  # noqa: N802
        return 0 if parent.isValid() else len(self._rows)

    def columnCount(self, parent: QtCore.QModelIndex = QtCore.QModelIndex()) -> int:  # noqa: N802
        return 0 if parent.isValid() else len(self.COLUMNS)

    def headerData(self, section: int, orientation: QtCore.Qt.Orientation, role: int = QtCore.Qt.DisplayRole):  # noqa: N802
        if role != QtCore.Qt.DisplayRole:
            return None
        if orientation == QtCore.Qt.Horizontal:
            if 0 <= section < len(self.COLUMNS):
                return self.COLUMNS[section]
        return None

    def data(self, index: QtCore.QModelIndex, role: int = QtCore.Qt.DisplayRole):  # noqa: N802
        if not index.isValid() or not (0 <= index.row() < len(self._rows)):
            return None
        r = self._rows[index.row()]
        col = index.column()

        if role in (QtCore.Qt.DisplayRole, QtCore.Qt.EditRole):
            sid = str(r.get("student_id", "") or "")
            name = f"{r.get('first_name','') or ''} {r.get('last_name','') or ''}".strip()
            age = r.get("age", "")
            cls = str(r.get("class", "") or "")
            sec = str(r.get("section", "") or "")
            p1 = str(r.get("primary_contact", "") or "")
            p2 = str(r.get("secondary_contact", "") or "")
            vals = [sid, name, "" if age is None else str(age), cls, sec, p1, p2]
            if 0 <= col < len(vals):
                return vals[col]
        if role == QtCore.Qt.UserRole:
            return r
        return None

    def row_dict(self, row: int) -> dict[str, Any] | None:
        if 0 <= row < len(self._rows):
            return self._rows[row]
        return None


class StudentFilterProxyModel(QtCore.QSortFilterProxyModel):
    def __init__(self):
        super().__init__()
        self.search_text = ""
        self.field = "All"
        self.class_filter = "(All classes)"
        self.section_filter = "(All sections)"

    def set_filters(self, *, search: str, field: str, cls: str, sec: str) -> None:
        self.search_text = (search or "").strip().lower()
        self.field = field or "All"
        self.class_filter = cls or "(All classes)"
        self.section_filter = sec or "(All sections)"
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row: int, source_parent: QtCore.QModelIndex) -> bool:  # noqa: N802
        model = self.sourceModel()
        if model is None:
            return True
        idx = model.index(source_row, 0, source_parent)
        row = model.data(idx, QtCore.Qt.UserRole)
        if not isinstance(row, dict):
            return True

        if self.class_filter != "(All classes)":
            if str(row.get("class", "") or "").strip() != self.class_filter:
                return False
        if self.section_filter != "(All sections)":
            if str(row.get("section", "") or "").strip() != self.section_filter:
                return False

        q = self.search_text
        if not q:
            return True

        name = f"{row.get('first_name','') or ''} {row.get('last_name','') or ''}".strip()
        blob_map = {
            "All": " ".join(
                [
                    str(row.get("student_id", "") or ""),
                    name,
                    str(row.get("age", "") or ""),
                    str(row.get("class", "") or ""),
                    str(row.get("section", "") or ""),
                    str(row.get("primary_contact", "") or ""),
                    str(row.get("secondary_contact", "") or ""),
                ]
            ),
            "Name": name,
            "ID": str(row.get("student_id", "") or ""),
            "Age": str(row.get("age", "") or ""),
            "Class": str(row.get("class", "") or ""),
            "Section": str(row.get("section", "") or ""),
            "Contact": f"{row.get('primary_contact','') or ''} {row.get('secondary_contact','') or ''}",
        }
        blob = str(blob_map.get(self.field, blob_map["All"]))
        return q in blob.lower()


class StudentDialog(QtWidgets.QDialog):
    def __init__(
        self,
        parent: QtWidgets.QWidget,
        *,
        title: str,
        student_id: str,
        initial: dict[str, Any] | None,
        custom_fields: list[str],
    ):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.setMinimumWidth(520)

        self.student_id = student_id
        self.initial = initial or {}
        self.custom_fields = custom_fields

        root = QtWidgets.QVBoxLayout(self)

        hdr = QtWidgets.QLabel(title)
        hdr.setObjectName("TopTitle")
        root.addWidget(hdr)

        root.addWidget(QtWidgets.QLabel(f"Student ID: {student_id}"))

        form = QtWidgets.QFormLayout()
        self.first = QtWidgets.QLineEdit(str(self.initial.get("first_name", "") or ""))
        self.last = QtWidgets.QLineEdit(str(self.initial.get("last_name", "") or ""))

        self.age = QtWidgets.QSpinBox()
        self.age.setRange(0, 120)
        try:
            self.age.setValue(int(str(self.initial.get("age", "") or "0").strip() or 0))
        except Exception:
            self.age.setValue(0)

        self.cls = QtWidgets.QLineEdit(str(self.initial.get("class", "") or ""))
        self.sec = QtWidgets.QLineEdit(str(self.initial.get("section", "") or ""))
        self.p1 = QtWidgets.QLineEdit(str(self.initial.get("primary_contact", "") or ""))
        self.p2 = QtWidgets.QLineEdit(str(self.initial.get("secondary_contact", "") or ""))

        form.addRow("First name *", self.first)
        form.addRow("Last name", self.last)
        form.addRow("Age (0 = unknown)", self.age)
        form.addRow("Class", self.cls)
        form.addRow("Section", self.sec)
        form.addRow("Primary contact", self.p1)
        form.addRow("Secondary contact", self.p2)

        self._custom: dict[str, QtWidgets.QLineEdit] = {}
        if self.custom_fields:
            sep = QtWidgets.QLabel("Custom fields")
            sep.setObjectName("SectionTitle")
            root.addSpacing(6)
            root.addWidget(sep)
            for f in self.custom_fields:
                w = QtWidgets.QLineEdit(str(self.initial.get(f, "") or ""))
                self._custom[f] = w
                form.addRow(f, w)

        root.addLayout(form)

        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        btns.accepted.connect(self._on_ok)
        btns.rejected.connect(self.reject)
        root.addWidget(btns)

    def _on_ok(self) -> None:
        if not self.first.text().strip():
            QtWidgets.QMessageBox.warning(self, "Validation", "First name is required.")
            self.first.setFocus()
            return
        self.accept()

    def get_data(self) -> dict[str, Any]:
        age = int(self.age.value())
        d: dict[str, Any] = {
            "student_id": self.student_id,
            "first_name": self.first.text().strip(),
            "last_name": self.last.text().strip(),
            "age": "" if age <= 0 else age,
            "class": self.cls.text().strip(),
            "section": self.sec.text().strip(),
            "primary_contact": self.p1.text().strip(),
            "secondary_contact": self.p2.text().strip(),
        }
        for k, w in self._custom.items():
            d[k] = w.text().strip()
        return d


class Card(QtWidgets.QFrame):
    def __init__(self, title: str, value: str = "0", *, accent: str | None = None):
        super().__init__()
        self.setProperty("class", "Card")
        self.setObjectName("Card")
        self.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.setStyleSheet("QFrame#Card { }")
        self.setProperty("qssClass", "Card")
        self.setProperty("class", "Card")
        self.setObjectName("Card")

        lay = QtWidgets.QVBoxLayout(self)
        lay.setContentsMargins(16, 16, 16, 16)
        lay.setSpacing(6)

        self.lbl_value = QtWidgets.QLabel(value)
        self.lbl_value.setProperty("class", "CardValue")
        self.lbl_value.setObjectName("CardValue")

        self.lbl_title = QtWidgets.QLabel(title)
        self.lbl_title.setProperty("class", "CardLabel")
        self.lbl_title.setObjectName("CardLabel")

        # Apply simple accent underline.
        if accent:
            line = QtWidgets.QFrame()
            line.setFixedHeight(3)
            line.setStyleSheet(f"background:{accent}; border-radius:2px;")
            lay.addWidget(line)

        lay.addWidget(self.lbl_value)
        lay.addWidget(self.lbl_title)
        lay.addStretch(1)

    def set_value(self, v: str) -> None:
        self.lbl_value.setText(v)


class DashboardPage(QtWidgets.QWidget):
    def __init__(self, parent: QtWidgets.QWidget):
        super().__init__(parent)
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(14)

        self.cards_row = QtWidgets.QHBoxLayout()
        self.cards_row.setSpacing(14)

        self.card_students = QtWidgets.QFrame()
        self.card_students.setObjectName("Card")
        self.card_students.setProperty("class", "Card")
        self.card_teachers = QtWidgets.QFrame()
        self.card_teachers.setObjectName("Card")
        self.card_teachers.setProperty("class", "Card")
        self.card_paid = QtWidgets.QFrame()
        self.card_paid.setObjectName("Card")
        self.card_paid.setProperty("class", "Card")

        self._card_students = Card("Students", "0", accent="#ffb100")
        self._card_teachers = Card("Teachers", "0", accent="#60a5fa")
        self._card_paid = Card("Student payments (Paid)", "0", accent="#34d399")

        self.cards_row.addWidget(self._card_students, 1)
        self.cards_row.addWidget(self._card_teachers, 1)
        self.cards_row.addWidget(self._card_paid, 1)

        root.addLayout(self.cards_row)

        charts_row = QtWidgets.QHBoxLayout()
        charts_row.setSpacing(14)

        self.chart_payments = QChartView()
        self.chart_payments.setRenderHint(QtGui.QPainter.Antialiasing)
        self.chart_payments.setMinimumHeight(320)
        self.chart_payments.setStyleSheet("background: transparent;")

        self.chart_classes = QChartView()
        self.chart_classes.setRenderHint(QtGui.QPainter.Antialiasing)
        self.chart_classes.setMinimumHeight(320)
        self.chart_classes.setStyleSheet("background: transparent;")

        wrap1 = QtWidgets.QFrame(); wrap1.setObjectName("Card"); wrap1.setProperty("class", "Card")
        w1 = QtWidgets.QVBoxLayout(wrap1); w1.setContentsMargins(12, 12, 12, 12)
        w1.addWidget(QtWidgets.QLabel("Payments (Paid vs Pending)"))
        w1.addWidget(self.chart_payments)

        wrap2 = QtWidgets.QFrame(); wrap2.setObjectName("Card"); wrap2.setProperty("class", "Card")
        w2 = QtWidgets.QVBoxLayout(wrap2); w2.setContentsMargins(12, 12, 12, 12)
        w2.addWidget(QtWidgets.QLabel("Students by Class"))
        w2.addWidget(self.chart_classes)

        charts_row.addWidget(wrap1, 1)
        charts_row.addWidget(wrap2, 1)

        root.addLayout(charts_row, 1)

    def set_counts(self, *, students: int, teachers: int, student_paid: int) -> None:
        self._card_students.set_value(str(students))
        self._card_teachers.set_value(str(teachers))
        self._card_paid.set_value(str(student_paid))

    def set_payments_chart(self, *, paid: int, pending: int) -> None:
        series = QPieSeries()
        series.append("Paid", max(paid, 0))
        series.append("Pending", max(pending, 0))
        for sl in series.slices():
            sl.setLabelVisible(True)
            sl.setLabelColor(QtGui.QColor(236, 240, 244))
        # Accent colors
        if series.count() >= 2:
            series.slices()[0].setBrush(QtGui.QColor(52, 211, 153))
            series.slices()[1].setBrush(QtGui.QColor(96, 165, 250))

        chart = QChart()
        chart.addSeries(series)
        chart.setBackgroundVisible(False)
        chart.legend().setLabelColor(QtGui.QColor(236, 240, 244))
        chart.legend().setAlignment(QtCore.Qt.AlignBottom)
        self.chart_payments.setChart(chart)

    def set_classes_chart(self, class_counts: dict[str, int]) -> None:
        cats = list(class_counts.keys())
        vals = [class_counts[k] for k in cats]

        barset = QBarSet("Students")
        barset.append(vals)
        barset.setColor(QtGui.QColor(255, 177, 0))

        series = QBarSeries()
        series.append(barset)

        chart = QChart()
        chart.addSeries(series)
        chart.setBackgroundVisible(False)
        chart.legend().setLabelColor(QtGui.QColor(236, 240, 244))
        chart.legend().setVisible(False)

        axis_x = QBarCategoryAxis()
        axis_x.append(cats or ["-"])
        axis_x.setLabelsColor(QtGui.QColor(200, 206, 216))
        chart.addAxis(axis_x, QtCore.Qt.AlignBottom)
        series.attachAxis(axis_x)

        axis_y = QValueAxis()
        axis_y.setLabelFormat("%d")
        axis_y.setLabelsColor(QtGui.QColor(200, 206, 216))
        axis_y.setGridLineColor(QtGui.QColor(31, 36, 49))
        chart.addAxis(axis_y, QtCore.Qt.AlignLeft)
        series.attachAxis(axis_y)

        self.chart_classes.setChart(chart)


class TeacherTableModel(QtCore.QAbstractTableModel):
    COLUMNS = ["ID", "Name", "Role", "Primary", "Secondary"]

    def __init__(self, rows: list[dict[str, Any]] | None = None):
        super().__init__()
        self._rows: list[dict[str, Any]] = rows or []

    def set_rows(self, rows: list[dict[str, Any]]) -> None:
        self.beginResetModel()
        self._rows = rows
        self.endResetModel()

    def rowCount(self, parent: QtCore.QModelIndex = QtCore.QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self._rows)

    def columnCount(self, parent: QtCore.QModelIndex = QtCore.QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self.COLUMNS)

    def headerData(self, section: int, orientation: QtCore.Qt.Orientation, role: int = QtCore.Qt.DisplayRole):
        if role != QtCore.Qt.DisplayRole:
            return None
        if orientation == QtCore.Qt.Horizontal:
            if 0 <= section < len(self.COLUMNS):
                return self.COLUMNS[section]
        return None

    def data(self, index: QtCore.QModelIndex, role: int = QtCore.Qt.DisplayRole):
        if not index.isValid() or not (0 <= index.row() < len(self._rows)):
            return None
        r = self._rows[index.row()]
        col = index.column()

        if role in (QtCore.Qt.DisplayRole, QtCore.Qt.EditRole):
            tid = str(r.get("teacher_id", "") or "")
            name = f"{r.get('first_name','') or ''} {r.get('last_name','') or ''}".strip()
            role_val = str(r.get("role", "") or "")
            p1 = str(r.get("primary_contact", "") or "")
            p2 = str(r.get("secondary_contact", "") or "")
            vals = [tid, name, role_val, p1, p2]
            if 0 <= col < len(vals):
                return vals[col]
        if role == QtCore.Qt.UserRole:
            return r
        return None

    def row_dict(self, row: int) -> dict[str, Any] | None:
        if 0 <= row < len(self._rows):
            return self._rows[row]
        return None


class TeacherDialog(QtWidgets.QDialog):
    def __init__(
        self,
        parent: QtWidgets.QWidget,
        *,
        title: str,
        teacher_id: str,
        initial: dict[str, Any] | None,
        custom_fields: list[str],
    ):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.setMinimumWidth(520)

        self.teacher_id = teacher_id
        self.initial = initial or {}
        self.custom_fields = custom_fields

        root = QtWidgets.QVBoxLayout(self)

        hdr = QtWidgets.QLabel(title)
        hdr.setObjectName("TopTitle")
        root.addWidget(hdr)

        root.addWidget(QtWidgets.QLabel(f"Teacher ID: {teacher_id}"))

        form = QtWidgets.QFormLayout()
        self.first = QtWidgets.QLineEdit(str(self.initial.get("first_name", "") or ""))
        self.last = QtWidgets.QLineEdit(str(self.initial.get("last_name", "") or ""))
        self.role = QtWidgets.QLineEdit(str(self.initial.get("role", "") or ""))
        self.p1 = QtWidgets.QLineEdit(str(self.initial.get("primary_contact", "") or ""))
        self.p2 = QtWidgets.QLineEdit(str(self.initial.get("secondary_contact", "") or ""))

        form.addRow("First name *", self.first)
        form.addRow("Last name", self.last)
        form.addRow("Role/Subject", self.role)
        form.addRow("Primary contact", self.p1)
        form.addRow("Secondary contact", self.p2)

        self._custom: dict[str, QtWidgets.QLineEdit] = {}
        if self.custom_fields:
            sep = QtWidgets.QLabel("Custom fields")
            sep.setObjectName("SectionTitle")
            root.addSpacing(6)
            root.addWidget(sep)
            for f in self.custom_fields:
                w = QtWidgets.QLineEdit(str(self.initial.get(f, "") or ""))
                self._custom[f] = w
                form.addRow(f, w)

        root.addLayout(form)

        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        btns.accepted.connect(self._on_ok)
        btns.rejected.connect(self.reject)
        root.addWidget(btns)

    def _on_ok(self) -> None:
        if not self.first.text().strip():
            QtWidgets.QMessageBox.warning(self, "Validation", "First name is required.")
            self.first.setFocus()
            return
        self.accept()

    def get_data(self) -> dict[str, Any]:
        d: dict[str, Any] = {
            "teacher_id": self.teacher_id,
            "first_name": self.first.text().strip(),
            "last_name": self.last.text().strip(),
            "role": self.role.text().strip(),
            "primary_contact": self.p1.text().strip(),
            "secondary_contact": self.p2.text().strip(),
        }
        for k, w in self._custom.items():
            d[k] = w.text().strip()
        return d


class StudentsPage(QtWidgets.QWidget):
    def __init__(self, parent: QtWidgets.QWidget):
        super().__init__(parent)
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(12)

        tools = QtWidgets.QFrame()
        tools.setObjectName("Card")
        tools.setProperty("class", "Card")
        tlay = QtWidgets.QHBoxLayout(tools)
        tlay.setContentsMargins(12, 12, 12, 12)
        tlay.setSpacing(10)

        self.search = QtWidgets.QLineEdit()
        self.search.setPlaceholderText("Search students (name, ID, class, section, age, contact)")

        self.field = QtWidgets.QComboBox()
        self.field.addItems(["All", "Name", "ID", "Class", "Section", "Age", "Contact"])

        self.cls = QtWidgets.QComboBox(); self.cls.addItems(["(All classes)"])
        self.sec = QtWidgets.QComboBox(); self.sec.addItems(["(All sections)"])

        self.btn_add = QtWidgets.QPushButton("Add Student")
        self.btn_add.setProperty("class", "Primary")
        self.btn_edit = QtWidgets.QPushButton("Edit")
        self.btn_del = QtWidgets.QPushButton("Delete")
        self.btn_del.setProperty("class", "Danger")

        tlay.addWidget(self.search, 2)
        tlay.addWidget(self.field)
        tlay.addWidget(self.cls)
        tlay.addWidget(self.sec)
        tlay.addStretch(1)
        tlay.addWidget(self.btn_add)
        tlay.addWidget(self.btn_edit)
        tlay.addWidget(self.btn_del)

        root.addWidget(tools)

        self.table = QtWidgets.QTableView()
        self.table.setSortingEnabled(True)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)

        root.addWidget(self.table, 1)


class TeachersPage(QtWidgets.QWidget):
    def __init__(self, parent: QtWidgets.QWidget):
        super().__init__(parent)
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(12)

        tools = QtWidgets.QFrame()
        tools.setObjectName("Card")
        tools.setProperty("class", "Card")
        tlay = QtWidgets.QHBoxLayout(tools)
        tlay.setContentsMargins(12, 12, 12, 12)
        tlay.setSpacing(10)

        self.search = QtWidgets.QLineEdit()
        self.search.setPlaceholderText("Search teachers (name, ID, role, contact)")

        self.btn_add = QtWidgets.QPushButton("Add Teacher")
        self.btn_add.setProperty("class", "Primary")
        self.btn_edit = QtWidgets.QPushButton("Edit")
        self.btn_del = QtWidgets.QPushButton("Delete")
        self.btn_del.setProperty("class", "Danger")

        tlay.addWidget(self.search, 2)
        tlay.addStretch(1)
        tlay.addWidget(self.btn_add)
        tlay.addWidget(self.btn_edit)
        tlay.addWidget(self.btn_del)

        root.addWidget(tools)

        self.table = QtWidgets.QTableView()
        self.table.setSortingEnabled(True)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)

        root.addWidget(self.table, 1)


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.err_logger = ErrorLogger()
        self.settings_store = SettingsStore()
        self.settings = self.settings_store.load()

        self.store = ExcelStore()
        self.store.ensure_workbook(self.settings.student_custom_fields, self.settings.teacher_custom_fields)

        self.selected = Selected()
        self._last_mtime: float = 0.0

        self.setWindowTitle(APP_NAME)
        self.resize(1360, 820)

        self._build_ui()
        self._wire()

        # Initial load
        self.refresh_all(rebuild_filters=True)

        # Poll file changes: this makes the table reflect external edits too.
        self._timer = QtCore.QTimer(self)
        self._timer.setInterval(800)
        self._timer.timeout.connect(self._tick)
        self._timer.start()

    # ---------- UI ----------
    def _build_ui(self) -> None:
        root = QtWidgets.QWidget()
        self.setCentralWidget(root)

        outer = QtWidgets.QHBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self.sidebar = QtWidgets.QFrame()
        self.sidebar.setObjectName("Sidebar")
        self.sidebar.setFixedWidth(240)
        sbl = QtWidgets.QVBoxLayout(self.sidebar)
        sbl.setContentsMargins(14, 14, 14, 14)
        sbl.setSpacing(10)

        title = QtWidgets.QLabel("High Tech\nSchool Management")
        title.setObjectName("AppTitle")
        sbl.addWidget(title)
        sbl.addSpacing(6)

        self.nav_buttons: dict[str, QtWidgets.QPushButton] = {}

        def nav_btn(text: str, key: str) -> QtWidgets.QPushButton:
            b = QtWidgets.QPushButton(text)
            b.setProperty("class", "NavBtn")
            b.setProperty("active", "false")
            b.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
            b.setMinimumHeight(40)
            self.nav_buttons[key] = b
            return b

        self.btn_dash = nav_btn("Dashboard", "dashboard")
        self.btn_students = nav_btn("Students", "students")
        self.btn_teachers = nav_btn("Teachers", "teachers")
        self.btn_payments = nav_btn("Payments", "payments")
        self.btn_activity = nav_btn("Activity Log", "activity")
        self.btn_settings = nav_btn("Settings", "settings")

        for b in [self.btn_dash, self.btn_students, self.btn_teachers, self.btn_payments, self.btn_activity, self.btn_settings]:
            sbl.addWidget(b)

        sbl.addStretch(1)
        sbl.addWidget(QtWidgets.QLabel("Data: school_data.xlsx\nLogs: error_log.txt"))

        # Content
        self.content = QtWidgets.QFrame()
        content_lay = QtWidgets.QVBoxLayout(self.content)
        content_lay.setContentsMargins(0, 0, 0, 0)
        content_lay.setSpacing(0)

        self.topbar = QtWidgets.QFrame()
        self.topbar.setObjectName("TopBar")
        tlay = QtWidgets.QHBoxLayout(self.topbar)
        tlay.setContentsMargins(16, 12, 16, 12)
        tlay.setSpacing(10)

        self.top_title = QtWidgets.QLabel("Dashboard")
        self.top_title.setObjectName("TopTitle")

        tlay.addWidget(self.top_title)
        tlay.addStretch(1)

        # Right-side quick actions
        self.act1 = QtWidgets.QPushButton("Action 1")
        self.act2 = QtWidgets.QPushButton("Action 2")
        self.act3 = QtWidgets.QPushButton("Action 3")
        self.act4 = QtWidgets.QPushButton("Action 4")
        for a in [self.act1, self.act2, self.act3, self.act4]:
            a.setMinimumHeight(34)
            tlay.addWidget(a)

        content_lay.addWidget(self.topbar)

        self.pages = QtWidgets.QStackedWidget()
        content_lay.addWidget(self.pages, 1)

        self.page_dashboard = DashboardPage(self.pages)
        self.page_students = StudentsPage(self.pages)
        self.page_teachers = TeachersPage(self.pages)
        self.page_payments = PaymentsPage(self.pages, self)
        self.page_activity = ActivityPage(self.pages)
        self.page_settings = SettingsPage(self.pages)

        self.pages.addWidget(self.page_dashboard)
        self.pages.addWidget(self.page_students)
        self.pages.addWidget(self.page_teachers)
        self.pages.addWidget(self.page_payments)
        self.pages.addWidget(self.page_activity)
        self.pages.addWidget(self.page_settings)

        outer.addWidget(self.sidebar)
        outer.addWidget(self.content, 1)

        # Students data model
        self.students_model = StudentTableModel([])
        self.students_proxy = StudentFilterProxyModel()
        self.students_proxy.setSourceModel(self.students_model)
        self.page_students.table.setModel(self.students_proxy)

        # Teachers data model
        self.teachers_model = TeacherTableModel([])
        self.teachers_proxy = QtCore.QSortFilterProxyModel()
        self.teachers_proxy.setSourceModel(self.teachers_model)
        self.page_teachers.table.setModel(self.teachers_proxy)

        self.statusBar().showMessage("Ready")

        # Apply card class styling by setting objectName for QSS
        for w in self.findChildren(QtWidgets.QFrame):
            if w.property("class") == "Card":
                w.setProperty("class", "Card")

    def _wire(self) -> None:
        self.btn_dash.clicked.connect(lambda: self.show_page("dashboard"))
        self.btn_students.clicked.connect(lambda: self.show_page("students"))
        self.btn_teachers.clicked.connect(lambda: self.show_page("teachers"))
        self.btn_payments.clicked.connect(lambda: self.show_page("payments"))
        self.btn_activity.clicked.connect(lambda: self.show_page("activity"))
        self.btn_settings.clicked.connect(lambda: self.show_page("settings"))

        # Students interactions
        self.page_students.search.textChanged.connect(self._apply_student_filters)
        self.page_students.field.currentTextChanged.connect(self._apply_student_filters)
        self.page_students.cls.currentTextChanged.connect(self._apply_student_filters)
        self.page_students.sec.currentTextChanged.connect(self._apply_student_filters)

        self.page_students.btn_add.clicked.connect(self.add_student)
        self.page_students.btn_edit.clicked.connect(self.edit_student)
        self.page_students.btn_del.clicked.connect(self.delete_student)

        sel = self.page_students.table.selectionModel()
        sel.selectionChanged.connect(self._on_student_selection_changed)

        # Teachers interactions
        self.page_teachers.search.textChanged.connect(self._apply_teacher_filter)
        self.page_teachers.btn_add.clicked.connect(self.add_teacher)
        self.page_teachers.btn_edit.clicked.connect(self.edit_teacher)
        self.page_teachers.btn_del.clicked.connect(self.delete_teacher)

        sel2 = self.page_teachers.table.selectionModel()
        sel2.selectionChanged.connect(self._on_teacher_selection_changed)

        # Payments interactions
        self.page_payments.entity_combo.currentTextChanged.connect(lambda: self.refresh_payments())
        self.page_payments.month_combo.currentIndexChanged.connect(lambda: self.refresh_payments())
        self.page_payments.year_spin.valueChanged.connect(lambda: self.refresh_payments())
        self.page_payments.filter_combo.currentTextChanged.connect(lambda: self.refresh_payments())
        self.page_payments.btn_set_payment.clicked.connect(self.set_payment)

        # Activity interactions
        self.page_activity.filter_action.currentTextChanged.connect(lambda: self.refresh_activity())
        self.page_activity.search.textChanged.connect(lambda: self.refresh_activity())

        # Settings interactions
        self.page_settings.btn_save.clicked.connect(self.save_settings)

    def _on_student_selection_changed(self, *_args) -> None:
        try:
            idxs = self.page_students.table.selectionModel().selectedRows()
            if not idxs:
                self.selected = Selected()
                return
            proxy_idx = idxs[0]
            src_idx = self.students_proxy.mapToSource(proxy_idx)
            row = self.students_model.row_dict(src_idx.row())
            sid = str((row or {}).get("student_id", "") or "").strip()
            if sid:
                self.selected = Selected(entity="student", person_id=sid)
        except Exception as e:
            self.err_logger.log_exception(e, "qt_student_selection")

    def _show_error(self, title: str, exc: BaseException) -> None:
        # Make failures visible to the user (Excel locks are very common on Windows).
        msg = str(exc)
        hint = ""
        if isinstance(exc, PermissionError) or "permission" in msg.lower() or "denied" in msg.lower():
            hint = (
                "\n\nHint: Close 'school_data.xlsx' in Excel/OneDrive preview and try again. "
                "Windows prevents saving while the file is open."
            )
        QtWidgets.QMessageBox.critical(self, title, f"{msg}{hint}")

    # ---------- Navigation ----------
    def show_page(self, key: str) -> None:
        mapping = {"dashboard": 0, "students": 1, "teachers": 2, "payments": 3, "activity": 4, "settings": 5}
        idx = mapping.get(key, 0)
        self.pages.setCurrentIndex(idx)
        self.top_title.setText({
            "dashboard": "Dashboard",
            "students": "Students",
            "teachers": "Teachers",
            "payments": "Payments",
            "activity": "Activity Log",
            "settings": "Settings",
        }.get(key, "Dashboard"))

        for k, b in self.nav_buttons.items():
            b.setProperty("active", "true" if k == key else "false")
            b.style().unpolish(b)
            b.style().polish(b)

        if key == "dashboard":
            self.refresh_dashboard()
        elif key == "students":
            self.refresh_students(rebuild_filters=False)
        elif key == "teachers":
            self.refresh_teachers()
        elif key == "payments":
            self.refresh_payments()
        elif key == "activity":
            self.refresh_activity()
        elif key == "settings":
            self.page_settings.load_settings(self.settings)

    # ---------- Live refresh ----------
    def _tick(self) -> None:
        try:
            p = Path(DATA_XLSX_PATH)
            if p.exists():
                m = p.stat().st_mtime
                if m != self._last_mtime:
                    self._last_mtime = m
                    # Reload fresh from disk
                    self.store.invalidate_cache()
                    self.refresh_all(rebuild_filters=True)
        except Exception as e:
            self.err_logger.log_exception(e, "qt_tick")

    # ---------- Data refresh ----------
    def refresh_all(self, *, rebuild_filters: bool) -> None:
        self.refresh_dashboard()
        self.refresh_students(rebuild_filters=rebuild_filters)
        self.refresh_teachers()

    def refresh_dashboard(self) -> None:
        try:
            students = self.store.list_students()
            teachers = self.store.list_teachers()

            year = int(self.settings.default_year)
            month = int(self.settings.default_month)
            stats = self.store.payment_stats("student", year, month)

            self.page_dashboard.set_counts(
                students=len(students),
                teachers=len(teachers),
                student_paid=int(stats.get("paid", 0)),
            )
            self.page_dashboard.set_payments_chart(
                paid=int(stats.get("paid", 0)),
                pending=int(stats.get("pending", 0)),
            )

            # Students by class
            cc: dict[str, int] = {}
            for r in students:
                c = str(r.get("class", "") or "").strip() or "(None)"
                cc[c] = cc.get(c, 0) + 1
            # Keep top 10 for a clean chart
            items = sorted(cc.items(), key=lambda kv: (-kv[1], kv[0]))[:10]
            self.page_dashboard.set_classes_chart({k: v for k, v in items})
        except Exception as e:
            self.err_logger.log_exception(e, "qt_refresh_dashboard")

    def refresh_students(self, *, rebuild_filters: bool) -> None:
        try:
            keep_id = self.selected.person_id if self.selected.entity == "student" else ""

            rows = self.store.list_students()
            rows.sort(key=lambda r: str(r.get("student_id", "") or ""))
            self.students_model.set_rows(rows)

            if rebuild_filters:
                classes = sorted({str(r.get("class", "") or "").strip() for r in rows if str(r.get("class", "") or "").strip()})
                sections = sorted({str(r.get("section", "") or "").strip() for r in rows if str(r.get("section", "") or "").strip()})

                self._refill_combo(self.page_students.cls, "(All classes)", classes)
                self._refill_combo(self.page_students.sec, "(All sections)", sections)

            self._apply_student_filters()

            if keep_id:
                self._select_student_by_id(keep_id)

            self.statusBar().showMessage(f"Students loaded: {len(rows)}")
        except Exception as e:
            self.err_logger.log_exception(e, "qt_refresh_students")

    def _refill_combo(self, combo: QtWidgets.QComboBox, first: str, values: list[str]) -> None:
        cur = combo.currentText()
        combo.blockSignals(True)
        combo.clear()
        combo.addItem(first)
        for v in values:
            combo.addItem(v)
        if cur and (cur == first or cur in values):
            combo.setCurrentText(cur)
        else:
            combo.setCurrentIndex(0)
        combo.blockSignals(False)

    def _apply_student_filters(self) -> None:
        self.students_proxy.set_filters(
            search=self.page_students.search.text(),
            field=self.page_students.field.currentText(),
            cls=self.page_students.cls.currentText(),
            sec=self.page_students.sec.currentText(),
        )

    def _select_student_by_id(self, student_id: str) -> None:
        # Search in source rows then map to proxy index.
        for r in range(self.students_model.rowCount()):
            row = self.students_model.row_dict(r) or {}
            if str(row.get("student_id", "") or "").strip() == student_id:
                src = self.students_model.index(r, 0)
                proxy = self.students_proxy.mapFromSource(src)
                if proxy.isValid():
                    self.page_students.table.selectRow(proxy.row())
                    self.page_students.table.scrollTo(proxy)
                break

    # ---------- CRUD ----------
    def _next_id(self, prefix: str, existing_ids: list[str]) -> str:
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

    def _emit(self, action: str, entity_type: str, entity_id: str, details: str = "") -> None:
        self.store.add_event(AppEvent(timestamp=now_ts(), action=action, entity_type=entity_type, entity_id=entity_id, details=details))

    def add_student(self) -> None:
        try:
            existing = [str(s.get("student_id", "") or "") for s in self.store.list_students()]
            sid = self._next_id(self.settings.student_id_prefix, existing)

            dlg = StudentDialog(
                self,
                title="Add Student",
                student_id=sid,
                initial=None,
                custom_fields=self.settings.student_custom_fields,
            )
            if dlg.exec() != QtWidgets.QDialog.Accepted:
                return
            data = dlg.get_data()

            self.store.upsert_student(data)
            self._emit("add_student", "student", sid, f"{data.get('first_name','')} {data.get('last_name','')}".strip())
            # Force UI to reflect what's actually in the xlsx on disk.
            self.store.invalidate_cache()
            self.refresh_all(rebuild_filters=True)
            self.selected = Selected(entity="student", person_id=sid)
            self._select_student_by_id(sid)
        except Exception as e:
            self.err_logger.log_exception(e, "qt_add_student")
            self._show_error("Add student failed", e)

    def _get_student(self, sid: str) -> dict[str, Any] | None:
        for r in self.store.list_students():
            if str(r.get("student_id", "") or "").strip() == sid:
                return r
        return None

    def edit_student(self) -> None:
        try:
            if self.selected.entity != "student" or not self.selected.person_id:
                return
            sid = self.selected.person_id
            current = self._get_student(sid) or {"student_id": sid}

            dlg = StudentDialog(
                self,
                title="Edit Student",
                student_id=sid,
                initial=current,
                custom_fields=self.settings.student_custom_fields,
            )
            if dlg.exec() != QtWidgets.QDialog.Accepted:
                return
            data = dlg.get_data()

            self.store.upsert_student(data)
            self._emit("edit_student", "student", sid, "updated")
            self.store.invalidate_cache()
            self.refresh_all(rebuild_filters=True)
            self.selected = Selected(entity="student", person_id=sid)
            self._select_student_by_id(sid)
        except Exception as e:
            self.err_logger.log_exception(e, "qt_edit_student")
            self._show_error("Edit student failed", e)

    def delete_student(self) -> None:
        try:
            if self.selected.entity != "student" or not self.selected.person_id:
                return
            sid = self.selected.person_id
            if QtWidgets.QMessageBox.question(self, "Confirm", f"Delete student {sid}?") != QtWidgets.QMessageBox.Yes:
                return
            ok = self.store.delete_student(sid)
            if ok:
                self._emit("delete_student", "student", sid, "deleted")
            self.selected = Selected()
            self.store.invalidate_cache()
            self.refresh_all(rebuild_filters=True)
        except Exception as e:
            self.err_logger.log_exception(e, "qt_delete_student")
            self._show_error("Delete student failed", e)

    # ---------- Teachers CRUD ----------
    def refresh_teachers(self) -> None:
        try:
            keep_id = self.selected.person_id if self.selected.entity == "teacher" else ""

            rows = self.store.list_teachers()
            rows.sort(key=lambda r: str(r.get("teacher_id", "") or ""))
            self.teachers_model.set_rows(rows)

            self._apply_teacher_filter()

            if keep_id:
                self._select_teacher_by_id(keep_id)

            self.statusBar().showMessage(f"Teachers loaded: {len(rows)}")
        except Exception as e:
            self.err_logger.log_exception(e, "qt_refresh_teachers")

    def _apply_teacher_filter(self) -> None:
        search = self.page_teachers.search.text().lower()
        for row in range(self.teachers_proxy.sourceModel().rowCount()):
            show = True
            if search:
                row_dict = self.teachers_model.row_dict(row) or {}
                blob = " ".join(str(v or "") for v in row_dict.values()).lower()
                if search not in blob:
                    show = False
            self.page_teachers.table.setRowHidden(row, not show)

    def _select_teacher_by_id(self, teacher_id: str) -> None:
        for r in range(self.teachers_model.rowCount()):
            row = self.teachers_model.row_dict(r) or {}
            if str(row.get("teacher_id", "") or "").strip() == teacher_id:
                self.page_teachers.table.selectRow(r)
                self.page_teachers.table.scrollToBottom()
                break

    def _on_teacher_selection_changed(self, *_args) -> None:
        try:
            idxs = self.page_teachers.table.selectionModel().selectedRows()
            if not idxs:
                return
            idx = idxs[0]
            row = self.teachers_model.row_dict(idx.row())
            tid = str((row or {}).get("teacher_id", "") or "").strip()
            if tid:
                self.selected = Selected(entity="teacher", person_id=tid)
        except Exception as e:
            self.err_logger.log_exception(e, "qt_teacher_selection")

    def _get_teacher(self, tid: str) -> dict[str, Any] | None:
        for r in self.store.list_teachers():
            if str(r.get("teacher_id", "") or "").strip() == tid:
                return r
        return None

    def add_teacher(self) -> None:
        try:
            existing = [str(t.get("teacher_id", "") or "") for t in self.store.list_teachers()]
            tid = self._next_id(self.settings.teacher_id_prefix, existing)

            dlg = TeacherDialog(
                self,
                title="Add Teacher",
                teacher_id=tid,
                initial=None,
                custom_fields=self.settings.teacher_custom_fields,
            )
            if dlg.exec() != QtWidgets.QDialog.Accepted:
                return
            data = dlg.get_data()

            self.store.upsert_teacher(data)
            self._emit("add_teacher", "teacher", tid, f"{data.get('first_name','')} {data.get('last_name','')}".strip())
            self.store.invalidate_cache()
            self.refresh_all(rebuild_filters=True)
            self.selected = Selected(entity="teacher", person_id=tid)
            self._select_teacher_by_id(tid)
        except Exception as e:
            self.err_logger.log_exception(e, "qt_add_teacher")
            self._show_error("Add teacher failed", e)

    def edit_teacher(self) -> None:
        try:
            if self.selected.entity != "teacher" or not self.selected.person_id:
                return
            tid = self.selected.person_id
            current = self._get_teacher(tid) or {"teacher_id": tid}

            dlg = TeacherDialog(
                self,
                title="Edit Teacher",
                teacher_id=tid,
                initial=current,
                custom_fields=self.settings.teacher_custom_fields,
            )
            if dlg.exec() != QtWidgets.QDialog.Accepted:
                return
            data = dlg.get_data()

            self.store.upsert_teacher(data)
            self._emit("edit_teacher", "teacher", tid, "updated")
            self.store.invalidate_cache()
            self.refresh_all(rebuild_filters=True)
            self.selected = Selected(entity="teacher", person_id=tid)
            self._select_teacher_by_id(tid)
        except Exception as e:
            self.err_logger.log_exception(e, "qt_edit_teacher")
            self._show_error("Edit teacher failed", e)

    def delete_teacher(self) -> None:
        try:
            if self.selected.entity != "teacher" or not self.selected.person_id:
                return
            tid = self.selected.person_id
            if QtWidgets.QMessageBox.question(self, "Confirm", f"Delete teacher {tid}?") != QtWidgets.QMessageBox.Yes:
                return
            ok = self.store.delete_teacher(tid)
            if ok:
                self._emit("delete_teacher", "teacher", tid, "deleted")
            self.selected = Selected()
            self.store.invalidate_cache()
            self.refresh_all(rebuild_filters=True)
        except Exception as e:
            self.err_logger.log_exception(e, "qt_delete_teacher")
            self._show_error("Delete teacher failed", e)

    # ---------- Payments ----------
    def refresh_payments(self) -> None:
        self.page_payments.refresh_payments(self.store, self.settings)

    def set_payment(self) -> None:
        try:
            result = self.page_payments.get_selected_person()
            if not result:
                QtWidgets.QMessageBox.warning(self, "No selection", "Please select a person from the table first.")
                return
            person_id, name = result

            entity_text = self.page_payments.entity_combo.currentText().lower().rstrip("s")
            month = self.page_payments.month_combo.currentIndex() + 1
            year = int(self.page_payments.year_spin.value())

            default_amt = self.settings.default_student_fee if entity_text == "student" else self.settings.default_teacher_salary

            rec = self.store.get_payment_record(entity_text, person_id, year, month)
            current_status = str(rec.get("status", "Pending") or "Pending")
            current_amount = float(rec.get("amount", 0) or 0)

            dlg = PaymentDialog(
                self,
                entity=entity_text,
                person_id=person_id,
                name=name,
                year=year,
                month=month,
                current_status=current_status,
                current_amount=current_amount,
                default_amount=default_amt,
            )
            if dlg.exec() != QtWidgets.QDialog.Accepted:
                return
            data = dlg.get_data()

            self.store.set_payment(entity_text, person_id, year, month, data["status"], data["amount"])
            self._emit("set_payment", entity_text, person_id, f"{MONTHS[month-1]} {year}: {data['status']} ${data['amount']}")
            self.store.invalidate_cache()
            self.refresh_payments()
        except Exception as e:
            self.err_logger.log_exception(e, "qt_set_payment")
            self._show_error("Set payment failed", e)

    # ---------- Activity ----------
    def refresh_activity(self) -> None:
        self.page_activity.refresh_activity(self.store)

    # ---------- Settings ----------
    def save_settings(self) -> None:
        try:
            new_settings = self.page_settings.get_settings()
            old_custom_students = set(self.settings.student_custom_fields)
            old_custom_teachers = set(self.settings.teacher_custom_fields)
            new_custom_students = set(new_settings.student_custom_fields)
            new_custom_teachers = set(new_settings.teacher_custom_fields)

            self.settings_store.save(new_settings)
            self.settings = new_settings

            # If custom fields changed, ensure workbook columns exist
            if old_custom_students != new_custom_students or old_custom_teachers != new_custom_teachers:
                self.store.ensure_workbook()

            QtWidgets.QMessageBox.information(self, "Success", "Settings saved successfully!")
            self._emit("update_settings", "settings", "", "saved")
            self.store.invalidate_cache()
            self.refresh_all(rebuild_filters=True)
        except Exception as e:
            self.err_logger.log_exception(e, "qt_save_settings")
            self._show_error("Save settings failed", e)


def run_qt_app() -> None:
    # High DPI: crisp fonts/icons on Windows.
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)

    app = QtWidgets.QApplication([])
    app.setApplicationName(APP_NAME)
    app.setStyle("Fusion")
    app.setPalette(_app_dark_palette())
    app.setStyleSheet(_qss())

    w = MainWindow()
    w.show()
    app.exec()
