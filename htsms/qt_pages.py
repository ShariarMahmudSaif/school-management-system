"""Additional Qt page widgets for payments, activity, and settings."""
from __future__ import annotations

from datetime import datetime
from typing import Any

from PySide6 import QtCore, QtGui, QtWidgets

from .logger import ErrorLogger
from .settings_store import Settings
from .storage import ExcelStore

MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]


class PaymentDialog(QtWidgets.QDialog):
    """Dialog for setting payment status and amount for a person/month."""
    def __init__(self, parent: QtWidgets.QWidget, *, entity: str, person_id: str, name: str, year: int, month: int, current_status: str, current_amount: float, default_amount: float):
        super().__init__(parent)
        self.setWindowTitle(f"Set Payment: {name}")
        self.setModal(True)
        self.setMinimumWidth(420)

        self.entity = entity
        self.person_id = person_id
        self.year = year
        self.month = month

        root = QtWidgets.QVBoxLayout(self)

        hdr = QtWidgets.QLabel(f"Payment for {name}")
        hdr.setObjectName("TopTitle")
        root.addWidget(hdr)

        root.addWidget(QtWidgets.QLabel(f"{MONTHS[month-1]} {year}"))

        form = QtWidgets.QFormLayout()

        self.status = QtWidgets.QComboBox()
        self.status.addItems(["Paid", "Pending"])
        self.status.setCurrentText(current_status if current_status in ["Paid", "Pending"] else "Pending")

        self.amount = QtWidgets.QDoubleSpinBox()
        self.amount.setRange(0, 1000000)
        self.amount.setDecimals(2)
        self.amount.setSuffix(" USD")
        amt = current_amount if current_amount > 0 else default_amount
        self.amount.setValue(amt)

        form.addRow("Status", self.status)
        form.addRow("Amount", self.amount)

        root.addLayout(form)

        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        root.addWidget(btns)

    def get_data(self) -> dict[str, Any]:
        return {
            "status": self.status.currentText(),
            "amount": self.amount.value(),
        }


class PaymentsPage(QtWidgets.QWidget):
    """Comprehensive payment tracking: student fees and teacher salaries with filtering, rollover display."""
    def __init__(self, parent: QtWidgets.QWidget, main_window):
        super().__init__(parent)
        self.main_window = main_window

        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(12)

        # Top controls
        controls = QtWidgets.QFrame()
        controls.setObjectName("Card")
        cly = QtWidgets.QHBoxLayout(controls)
        cly.setContentsMargins(12, 12, 12, 12)
        cly.setSpacing(10)

        cly.addWidget(QtWidgets.QLabel("Entity:"))
        self.entity_combo = QtWidgets.QComboBox()
        self.entity_combo.addItems(["Students", "Teachers"])
        cly.addWidget(self.entity_combo)

        cly.addWidget(QtWidgets.QLabel("Month:"))
        self.month_combo = QtWidgets.QComboBox()
        self.month_combo.addItems(MONTHS)
        cly.addWidget(self.month_combo)

        cly.addWidget(QtWidgets.QLabel("Year:"))
        self.year_spin = QtWidgets.QSpinBox()
        self.year_spin.setRange(2020, 2030)
        self.year_spin.setValue(datetime.now().year)
        cly.addWidget(self.year_spin)

        cly.addWidget(QtWidgets.QLabel("Filter:"))
        self.filter_combo = QtWidgets.QComboBox()
        self.filter_combo.addItems(["All", "Paid", "Pending"])
        cly.addWidget(self.filter_combo)

        self.btn_set_payment = QtWidgets.QPushButton("Set Payment")
        self.btn_set_payment.setProperty("class", "Primary")
        cly.addWidget(self.btn_set_payment)

        cly.addStretch(1)
        root.addWidget(controls)

        # Summary card
        summary_card = QtWidgets.QFrame()
        summary_card.setObjectName("Card")
        sum_ly = QtWidgets.QHBoxLayout(summary_card)
        sum_ly.setContentsMargins(12, 12, 12, 12)
        self.lbl_summary = QtWidgets.QLabel("Select month/year and entity to view payments")
        self.lbl_summary.setObjectName("CardLabel")
        sum_ly.addWidget(self.lbl_summary)
        root.addWidget(summary_card)

        # Table
        self.table = QtWidgets.QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["ID", "Name", "Status", "Amount", "Pending Total", "Pending Months"])
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        root.addWidget(self.table, 1)

    def refresh_payments(self, store: ExcelStore, settings: Settings) -> None:
        """Load payment data for selected month/year/entity with status filtering."""
        try:
            entity_text = self.entity_combo.currentText().lower().rstrip("s")  # "Students" -> "student"
            month = self.month_combo.currentIndex() + 1
            year = int(self.year_spin.value())
            filter_status = self.filter_combo.currentText()

            default_amt = settings.default_student_fee if entity_text == "student" else settings.default_teacher_salary

            people = store.list_students() if entity_text == "student" else store.list_teachers()
            id_key = "student_id" if entity_text == "student" else "teacher_id"

            # Build table data
            table_data: list[dict[str, Any]] = []
            total_paid_count = 0
            total_pending_count = 0
            total_paid_amount = 0.0
            total_pending_amount = 0.0

            for p in people:
                pid = str(p.get(id_key, "") or "")
                if not pid:
                    continue

                name = f"{p.get('first_name','') or ''} {p.get('last_name','') or ''}".strip()

                # Get current month payment
                rec = store.get_payment_record(entity_text, pid, year, month)
                status = str(rec.get("status", "Pending") or "Pending")
                amount = float(rec.get("amount", 0) or 0)
                if amount <= 0:
                    amount = default_amt

                # Get total pending (all unpaid months up to now)
                pending_months = store.get_pending_months(entity_text, pid, year, month, default_amt)
                pending_total = sum(pm["amount"] for pm in pending_months)
                pending_months_str = ", ".join([f"{MONTHS[pm['month']-1][:3]} {pm['year']}" for pm in pending_months[-6:]])  # last 6
                if len(pending_months) > 6:
                    pending_months_str = f"...{pending_months_str}"

                # Apply filter
                if filter_status == "Paid" and status.lower() != "paid":
                    continue
                if filter_status == "Pending" and status.lower() == "paid":
                    continue

                table_data.append({
                    "id": pid,
                    "name": name,
                    "status": status,
                    "amount": amount,
                    "pending_total": pending_total,
                    "pending_months": pending_months_str,
                })

                if status.lower() == "paid":
                    total_paid_count += 1
                    total_paid_amount += amount
                else:
                    total_pending_count += 1
                    total_pending_amount += amount

            # Update summary
            entity_label = "Students" if entity_text == "student" else "Teachers"
            self.lbl_summary.setText(
                f"{entity_label} {MONTHS[month-1]} {year}: "
                f"Paid: {total_paid_count} (${total_paid_amount:,.2f}) | "
                f"Pending: {total_pending_count} (${total_pending_amount:,.2f})"
            )

            # Populate table
            self.table.setRowCount(len(table_data))
            for row_idx, row_data in enumerate(table_data):
                self.table.setItem(row_idx, 0, QtWidgets.QTableWidgetItem(row_data["id"]))
                self.table.setItem(row_idx, 1, QtWidgets.QTableWidgetItem(row_data["name"]))

                status_item = QtWidgets.QTableWidgetItem(row_data["status"])
                if row_data["status"].lower() == "paid":
                    status_item.setBackground(QtGui.QColor(34, 211, 153, 100))
                else:
                    status_item.setBackground(QtGui.QColor(225, 29, 72, 100))
                self.table.setItem(row_idx, 2, status_item)

                self.table.setItem(row_idx, 3, QtWidgets.QTableWidgetItem(f"${row_data['amount']:,.2f}"))
                self.table.setItem(row_idx, 4, QtWidgets.QTableWidgetItem(f"${row_data['pending_total']:,.2f}"))
                self.table.setItem(row_idx, 5, QtWidgets.QTableWidgetItem(row_data["pending_months"] or "-"))

            self.table.resizeColumnsToContents()
        except Exception as e:
            self.main_window.err_logger.log_exception(e, "qt_refresh_payments")

    def get_selected_person(self) -> tuple[str, str] | None:
        """Return (person_id, name) of selected row."""
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            return None
        row = rows[0].row()
        pid = self.table.item(row, 0).text() if self.table.item(row, 0) else ""
        name = self.table.item(row, 1).text() if self.table.item(row, 1) else ""
        if not pid:
            return None
        return (pid, name)


class ActivityPage(QtWidgets.QWidget):
    """Activity log with filtering."""
    def __init__(self, parent: QtWidgets.QWidget):
        super().__init__(parent)
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(12)

        controls = QtWidgets.QFrame()
        controls.setObjectName("Card")
        cly = QtWidgets.QHBoxLayout(controls)
        cly.setContentsMargins(12, 12, 12, 12)

        cly.addWidget(QtWidgets.QLabel("Filter action:"))
        self.filter_action = QtWidgets.QComboBox()
        self.filter_action.addItems(["All", "add_student", "edit_student", "delete_student", "add_teacher", "edit_teacher", "delete_teacher", "set_payment"])
        cly.addWidget(self.filter_action)

        self.search = QtWidgets.QLineEdit()
        self.search.setPlaceholderText("Search activity (entity ID, details)")
        cly.addWidget(self.search, 2)

        cly.addStretch(1)
        root.addWidget(controls)

        self.table = QtWidgets.QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["Timestamp", "Action", "Entity Type", "Entity ID", "Details"])
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        root.addWidget(self.table, 1)

    def refresh_activity(self, store: ExcelStore) -> None:
        try:
            events = store.list_events(limit=500)
            filter_action = self.filter_action.currentText()
            search_text = self.search.text().strip().lower()

            filtered: list[dict[str, Any]] = []
            for ev in reversed(events):
                if filter_action != "All" and ev.get("action", "") != filter_action:
                    continue
                if search_text:
                    blob = f"{ev.get('entity_id','')} {ev.get('details','')}".lower()
                    if search_text not in blob:
                        continue
                filtered.append(ev)

            self.table.setRowCount(len(filtered))
            for row_idx, ev in enumerate(filtered):
                self.table.setItem(row_idx, 0, QtWidgets.QTableWidgetItem(str(ev.get("timestamp", ""))))
                self.table.setItem(row_idx, 1, QtWidgets.QTableWidgetItem(str(ev.get("action", ""))))
                self.table.setItem(row_idx, 2, QtWidgets.QTableWidgetItem(str(ev.get("entity_type", ""))))
                self.table.setItem(row_idx, 3, QtWidgets.QTableWidgetItem(str(ev.get("entity_id", ""))))
                self.table.setItem(row_idx, 4, QtWidgets.QTableWidgetItem(str(ev.get("details", ""))))

            self.table.resizeColumnsToContents()
        except Exception as e:
            pass


class SettingsPage(QtWidgets.QWidget):
    """Settings editor: prefixes, custom fields, default fees/salaries, appearance."""
    def __init__(self, parent: QtWidgets.QWidget):
        super().__init__(parent)
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(12)

        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)

        scroll_widget = QtWidgets.QWidget()
        scroll_layout = QtWidgets.QVBoxLayout(scroll_widget)
        scroll_layout.setSpacing(16)

        # ID Prefixes
        id_card = QtWidgets.QFrame()
        id_card.setObjectName("Card")
        id_layout = QtWidgets.QVBoxLayout(id_card)
        id_layout.setContentsMargins(16, 16, 16, 16)
        lbl_id = QtWidgets.QLabel("ID Prefixes")
        lbl_id.setObjectName("TopTitle")
        id_layout.addWidget(lbl_id)

        id_form = QtWidgets.QFormLayout()
        self.student_prefix = QtWidgets.QLineEdit()
        self.teacher_prefix = QtWidgets.QLineEdit()
        id_form.addRow("Student ID prefix", self.student_prefix)
        id_form.addRow("Teacher ID prefix", self.teacher_prefix)
        id_layout.addLayout(id_form)
        scroll_layout.addWidget(id_card)

        # Default Amounts
        amt_card = QtWidgets.QFrame()
        amt_card.setObjectName("Card")
        amt_layout = QtWidgets.QVBoxLayout(amt_card)
        amt_layout.setContentsMargins(16, 16, 16, 16)
        lbl_amt = QtWidgets.QLabel("Default Payment Amounts")
        lbl_amt.setObjectName("TopTitle")
        amt_layout.addWidget(lbl_amt)

        amt_form = QtWidgets.QFormLayout()
        self.student_fee = QtWidgets.QDoubleSpinBox()
        self.student_fee.setRange(0, 100000)
        self.student_fee.setDecimals(2)
        self.student_fee.setSuffix(" USD")

        self.teacher_salary = QtWidgets.QDoubleSpinBox()
        self.teacher_salary.setRange(0, 100000)
        self.teacher_salary.setDecimals(2)
        self.teacher_salary.setSuffix(" USD")

        amt_form.addRow("Student monthly fee", self.student_fee)
        amt_form.addRow("Teacher monthly salary", self.teacher_salary)
        amt_layout.addLayout(amt_form)
        scroll_layout.addWidget(amt_card)

        # Custom Fields
        fields_card = QtWidgets.QFrame()
        fields_card.setObjectName("Card")
        fields_layout = QtWidgets.QVBoxLayout(fields_card)
        fields_layout.setContentsMargins(16, 16, 16, 16)
        lbl_fields = QtWidgets.QLabel("Custom Fields (comma-separated)")
        lbl_fields.setObjectName("TopTitle")
        fields_layout.addWidget(lbl_fields)

        fields_form = QtWidgets.QFormLayout()
        self.student_custom = QtWidgets.QLineEdit()
        self.teacher_custom = QtWidgets.QLineEdit()
        fields_form.addRow("Student fields", self.student_custom)
        fields_form.addRow("Teacher fields", self.teacher_custom)
        fields_layout.addLayout(fields_form)
        scroll_layout.addWidget(fields_card)

        # Default Month/Year
        date_card = QtWidgets.QFrame()
        date_card.setObjectName("Card")
        date_layout = QtWidgets.QVBoxLayout(date_card)
        date_layout.setContentsMargins(16, 16, 16, 16)
        lbl_date = QtWidgets.QLabel("Default Payment Period")
        lbl_date.setObjectName("TopTitle")
        date_layout.addWidget(lbl_date)

        date_form = QtWidgets.QFormLayout()
        self.default_month = QtWidgets.QComboBox()
        self.default_month.addItems(MONTHS)
        self.default_year = QtWidgets.QSpinBox()
        self.default_year.setRange(2020, 2030)
        date_form.addRow("Default month", self.default_month)
        date_form.addRow("Default year", self.default_year)
        date_layout.addLayout(date_form)
        scroll_layout.addWidget(date_card)

        # Save button
        self.btn_save = QtWidgets.QPushButton("Save Settings")
        self.btn_save.setProperty("class", "Primary")
        self.btn_save.setMinimumHeight(44)
        scroll_layout.addWidget(self.btn_save)

        scroll_layout.addStretch(1)
        scroll.setWidget(scroll_widget)
        root.addWidget(scroll)

    def load_settings(self, settings: Settings) -> None:
        self.student_prefix.setText(settings.student_id_prefix)
        self.teacher_prefix.setText(settings.teacher_id_prefix)
        self.student_fee.setValue(settings.default_student_fee)
        self.teacher_salary.setValue(settings.default_teacher_salary)
        self.student_custom.setText(", ".join(settings.student_custom_fields))
        self.teacher_custom.setText(", ".join(settings.teacher_custom_fields))
        self.default_month.setCurrentIndex(settings.default_month - 1)
        self.default_year.setValue(settings.default_year)

    def get_settings(self) -> Settings:
        def parse_fields(text: str) -> list[str]:
            return [f.strip() for f in text.split(",") if f.strip()]

        return Settings(
            student_id_prefix=self.student_prefix.text().strip(),
            teacher_id_prefix=self.teacher_prefix.text().strip(),
            student_custom_fields=parse_fields(self.student_custom.text()),
            teacher_custom_fields=parse_fields(self.teacher_custom.text()),
            default_year=int(self.default_year.value()),
            default_month=self.default_month.currentIndex() + 1,
            default_student_fee=float(self.student_fee.value()),
            default_teacher_salary=float(self.teacher_salary.value()),
        )
