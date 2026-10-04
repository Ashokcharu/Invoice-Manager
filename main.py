import sys
import os
import shutil
import sqlite3
import subprocess
import textwrap
import re
import json
import tempfile
from datetime import datetime
from copy import copy
from urllib.parse import quote

from pathlib import Path
from decimal import Decimal, ROUND_HALF_UP

from PySide6.QtCore import (
    QDate, QEvent, QTimer, Qt, QStandardPaths, QUrl, QRectF, QSize,
    QRegularExpression,
)
from PySide6.QtCore import QMimeData
from PySide6.QtGui import (
    QFont, QAction, QImage, QImageReader, QColor, QBrush, QDesktopServices,
    QPainter, QPen, QIcon, QPixmap, QRegularExpressionValidator,
)
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QFormLayout,
    QLabel,
    QLineEdit,
    QComboBox,
    QDoubleSpinBox,
    QAbstractSpinBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QMessageBox,
    QDateEdit,
    QHeaderView,
    QGroupBox,
    QToolButton,
    QScrollArea,
    QTabWidget,
    QFileDialog,
    QTextEdit,
    QCheckBox,
    QListWidget,
    QListWidgetItem,
    QDialog,
    QDialogButtonBox,
    QStackedWidget,
    QFrame,
    QStyle,
)

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.drawing.image import Image as SpreadsheetImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
import qrcode


# ============================================================
# PATHS
# ============================================================

APP_DIR = Path(__file__).resolve().parent
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", APP_DIR))

DATA_DIR = APP_DIR / "data"
INVOICE_DIR = APP_DIR / "invoices"
PRIVATE_INVOICE_DIR = INVOICE_DIR / "Private"

DB_PATH = DATA_DIR / "invoices.db"
SIGNATURE_PATH = DATA_DIR / "invoice_signature.png"
START_INVOICE_NUMBER = 101

DATA_DIR.mkdir(exist_ok=True)
INVOICE_DIR.mkdir(exist_ok=True)


# ============================================================
# MONEY
# ============================================================

def money(value):
    return Decimal(str(value)).quantize(
        Decimal("0.01"),
        rounding=ROUND_HALF_UP
    )


def normalize_indian_mobile(value):
    digits = re.sub(r"\D", "", value or "")
    if len(digits) == 11 and digits.startswith("0"):
        digits = digits[1:]
    if len(digits) == 12 and digits.startswith("91"):
        digits = digits[2:]
    if len(digits) != 10 or digits[0] not in "6789":
        return ""
    return f"91{digits}"


def number_to_words(value):
    value = int(value)
    if value == 0:
        return "Zero"

    ones = (
        "Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven",
        "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen",
        "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen",
        "Nineteen",
    )
    tens = (
        "", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty",
        "Seventy", "Eighty", "Ninety",
    )

    def under_thousand(number):
        words = []
        if number >= 100:
            words.extend((ones[number // 100], "Hundred"))
            number %= 100
        if number >= 20:
            words.append(tens[number // 10])
            number %= 10
        if number:
            words.append(ones[number])
        return " ".join(words)

    words = []
    for divisor, label in (
        (10**11, "Kharab"), (10**9, "Arab"), (10**7, "Crore"),
        (10**5, "Lakh"), (1000, "Thousand"), (100, "Hundred"),
    ):
        if value >= divisor:
            group, value = divmod(value, divisor)
            words.append(f"{under_thousand(group)} {label}")
    if value:
        words.append(under_thousand(value))
    return " ".join(words)


def amount_to_words(value):
    value = money(value)
    rupees = int(value)
    paise = int((value - rupees) * 100)
    words = f"Rupees {number_to_words(rupees)}"
    if paise:
        words += f" and {number_to_words(paise)} Paise"
    return f"{words} Only"


def invoice_filename(invoice_no, buyer_name, extension, document_kind="INVOICE"):
    buyer_part = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", buyer_name or "")
    buyer_part = re.sub(r"\s+", "_", buyer_part).strip(" ._")[:80]
    buyer_part = buyer_part.rstrip(" ._") or "Buyer"
    prefix = "QTN" if document_kind == "QUOTATION" else "INV"
    return f"{prefix}-{invoice_no:04d}_{buyer_part}{extension}"


def document_kind_for_selection(index):
    return "INVOICE" if index == 0 else "QUOTATION"


def document_kind_selection_for_kind(document_kind):
    return 1 if document_kind == "QUOTATION" else 0


def document_folder_for_selection(index):
    return "Invoices" if index == 0 else "Quotations"


class NoWheelDoubleSpinBox(QDoubleSpinBox):

    def wheelEvent(self, event):
        event.ignore()


class AmountEntryDoubleSpinBox(NoWheelDoubleSpinBox):

    def __init__(self):
        super().__init__()
        self.lineEdit().installEventFilter(self)

    def focusInEvent(self, event):
        super().focusInEvent(event)
        self.selectAll()

    def eventFilter(self, watched, event):
        if (
            watched is self.lineEdit()
            and event.type() == QEvent.Type.MouseButtonPress
        ):
            QTimer.singleShot(0, self.selectAll)
        return super().eventFilter(watched, event)


class NoWheelComboBox(QComboBox):

    def wheelEvent(self, event):
        event.ignore()


class NoWheelDateEdit(QDateEdit):

    def wheelEvent(self, event):
        event.ignore()


class FinanceDonutChart(QWidget):

    def __init__(self, parent=None):
        super().__init__(parent)
        self.collected = 0.0
        self.outstanding = 0.0
        self.setMinimumSize(220, 190)

    def set_values(self, collected, outstanding):
        self.collected = max(0.0, collected)
        self.outstanding = max(0.0, outstanding)
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        side = min(self.width() - 24, self.height() - 24)
        chart = QRectF(
            (self.width() - side) / 2,
            (self.height() - side) / 2,
            side,
            side,
        )
        total = self.collected + self.outstanding
        if total <= 0:
            painter.setPen(QPen(QColor("#d8e1e8"), 18))
            painter.drawArc(chart, 0, 360 * 16)
            return

        collected_angle = round(self.collected / total * 360 * 16)
        painter.setPen(QPen(QColor("#21866b"), 18))
        painter.drawArc(chart, 90 * 16, -collected_angle)
        painter.setPen(QPen(QColor("#e7a34d"), 18))
        painter.drawArc(chart, 90 * 16 - collected_angle, -(360 * 16 - collected_angle))
        painter.setPen(QColor("#253b4a"))
        painter.setFont(QFont("Segoe UI", 10, QFont.Weight.DemiBold))
        painter.drawText(chart, Qt.AlignmentFlag.AlignCenter, "INVOICED")


class FinanceMonthlyChart(QWidget):

    MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    def __init__(self, parent=None):
        super().__init__(parent)
        self.values = [0.0] * 12
        self.setMinimumHeight(250)

    def set_values(self, values):
        self.values = list(values[:12]) + [0.0] * max(0, 12 - len(values))
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        left, right, top, bottom = 42, 16, 16, 34
        chart = self.rect().adjusted(left, top, -right, -bottom)
        max_value = max(self.values, default=0)
        max_value = max_value or 1
        painter.setPen(QPen(QColor("#dce4e9"), 1))
        for fraction in (0, 0.5, 1):
            y = chart.bottom() - chart.height() * fraction
            painter.drawLine(chart.left(), round(y), chart.right(), round(y))
        slot = chart.width() / 12
        bar_width = max(7, slot * 0.54)
        for index, value in enumerate(self.values):
            bar_height = chart.height() * value / max_value
            bar = QRectF(
                chart.left() + slot * index + (slot - bar_width) / 2,
                chart.bottom() - bar_height,
                bar_width,
                bar_height,
            )
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor("#21866b" if value else "#e3e9ed"))
            painter.drawRoundedRect(bar, 3, 3)
            painter.setPen(QColor("#667783"))
            painter.setFont(QFont("Segoe UI", 8))
            painter.drawText(
                QRectF(chart.left() + slot * index, chart.bottom() + 7, slot, 20),
                Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop,
                self.MONTHS[index],
            )


# ============================================================
# APPLICATION
# ============================================================

class InvoiceApp(QMainWindow):

    def __init__(self):
        super().__init__()

        self.pending_private_move_paths = []
        self.private_move_cancelled = False
        self.active_document_kind = None
        self.setWindowTitle("Invoice Manager")
        self.setWindowIcon(QIcon(str(RESOURCE_DIR / "app.ico")))
        self.resize(1280, 820)
        self.setMinimumSize(1050, 700)

        self.init_database()

        self.windows_style()

        self.create_menu()

        self.create_ui()

        self.delete_document_draft("INVOICE")
        self.delete_document_draft("QUOTATION")
        self.new_invoice()


    # ========================================================
    # WINDOWS STYLE
    # ========================================================

    def windows_style(self):

        self.setStyleSheet("""

        QMainWindow, QWidget {
            background: #f4f7fb;
            color: #1f2937;
            font-family: "Segoe UI";
            font-size: 10pt;
        }

        QLineEdit, QTextEdit, QComboBox, QDateEdit, QDoubleSpinBox,
        QTableWidget, QListWidget {
            background: #ffffff;
            color: #1f2937;
            border: 1px solid #bdc9d8;
            border-radius: 4px;
            padding: 6px;
            selection-background-color: #dceeff;
            selection-color: #123b5d;
        }
        QScrollBar:vertical {
            background: #e8eef4;
            width: 12px;
            margin: 2px;
            border-radius: 6px;
        }
        QScrollBar::handle:vertical {
            background: #7893a8;
            min-height: 34px;
            border-radius: 6px;
        }
        QScrollBar::handle:vertical:hover {
            background: #4e718d;
        }
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
            height: 0;
        }
        QScrollBar::horizontal {
            background: #e8eef4;
            height: 12px;
            margin: 2px;
            border-radius: 6px;
        }
        QScrollBar::handle:horizontal {
            background: #7893a8;
            min-width: 34px;
            border-radius: 6px;
        }
        QScrollBar::handle:horizontal:hover {
            background: #4e718d;
        }
        QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
            width: 0;
        }
        QTableWidget {
            alternate-background-color: #f4f7fb;
            background-color: #ffffff;
            color: #1f2937;
        }
        QTableWidget::item {
            color: #1f2937;
        }
        QLineEdit:focus, QTextEdit:focus, QComboBox:focus, QDateEdit:focus,
        QDoubleSpinBox:focus, QTableWidget:focus, QListWidget:focus {
            border: 1px solid #0078d4;
        }
        QGroupBox {
            background: #ffffff;
            border: 1px solid #cbd6e2;
            border-radius: 6px;
            margin-top: 16px;
            padding: 14px 12px 10px;
            font-weight: 600;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 12px;
            padding: 0 6px;
            color: #173b57;
        }
        QCheckBox {
            spacing: 8px;
            padding: 6px 12px;
            border: 1px solid #8aa1b3;
            border-radius: 4px;
            font-weight: 600;
        }
        QCheckBox::indicator {
            width: 22px;
            height: 22px;
        }
        QCheckBox:checked {
            background: #e8f3fb;
            border-color: #36749e;
        }
        QHeaderView::section {
            background: #eaf1f8;
            color: #173b57;
            border: 1px solid #cbd6e2;
            padding: 7px 6px;
            font-weight: bold;
        }
        QPushButton {
            background: #ffffff;
            color: #1f2937;
            border: 1px solid #aebdcd;
            border-radius: 4px;
            padding: 7px 14px;
        }
        QPushButton:hover {
            background: #edf5fc;
            border-color: #6d9fc7;
        }
        QPushButton:pressed {
            background: #dceeff;
        }
        QPushButton#primary {
            background: #21866b;
            border-color: #21866b;
            color: #ffffff;
            font-weight: 600;
        }
        QPushButton#primary:hover {
            background: #176b56;
            border-color: #176b56;
        }
        QListWidget::item {
            padding: 5px 8px;
            border-bottom: 1px solid #edf1f5;
        }
        QListWidget::item:selected {
            background: #dceeff;
            color: #123b5d;
        }
        """)


    def spinbox_with_arrows(self, spinbox, step_size=1):

        spinbox.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        container = QWidget()
        container_layout = QHBoxLayout(container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        container_layout.setSpacing(0)
        container_layout.addWidget(spinbox, 1)

        buttons = QWidget()
        buttons_layout = QVBoxLayout(buttons)
        buttons_layout.setContentsMargins(0, 0, 0, 0)
        buttons_layout.setSpacing(0)
        for arrow_type, step, tooltip in (
            (Qt.ArrowType.UpArrow, step_size, "Increase"),
            (Qt.ArrowType.DownArrow, -step_size, "Decrease"),
        ):
            button = QToolButton()
            button.setArrowType(arrow_type)
            button.setToolTip(tooltip)
            button.setFixedSize(22, 16)
            button.setStyleSheet(
                "QToolButton { background: #eaf1f8; color: #173b57; "
                "border: 1px solid #bdc9d8; padding: 0; }"
                "QToolButton:hover { background: #dceeff; }"
            )
            button.clicked.connect(
                lambda checked=False, amount=step: spinbox.stepBy(amount)
            )
            buttons_layout.addWidget(button)
        container_layout.addWidget(buttons)
        return container


    def create_menu(self):

        file_menu = self.menuBar().addMenu("File")
        new_action = QAction("New Invoice", self)
        new_action.triggered.connect(self.new_invoice)
        file_menu.addAction(new_action)
        save_action = QAction("Save XLSX and Download PDF", self)
        save_action.triggered.connect(
            lambda checked=False: self.save_document_bundle()
        )
        file_menu.addAction(save_action)
        file_menu.addSeparator()
        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)


    # ========================================================
    # DATABASE
    # ========================================================

    def init_database(self):

        with sqlite3.connect(DB_PATH) as con:

            con.execute("""
            
                CREATE TABLE IF NOT EXISTS invoices (

                    id INTEGER PRIMARY KEY AUTOINCREMENT,

                    invoice_no INTEGER UNIQUE NOT NULL,

                    invoice_date TEXT NOT NULL,

                    invoice_type TEXT NOT NULL,

                    customer_name TEXT,

                    grand_total REAL,

                    xlsx_path TEXT,

                    pdf_path TEXT,

                    created_at
                    TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )

            """)
            columns = {
                row[1] for row in con.execute("PRAGMA table_info(invoices)")
            }
            if "document_kind" not in columns:
                con.execute(
                    "ALTER TABLE invoices ADD COLUMN document_kind TEXT NOT NULL DEFAULT 'INVOICE'"
                )
            if "qr_scanner_path" not in columns:
                con.execute(
                    "ALTER TABLE invoices ADD COLUMN qr_scanner_path TEXT"
                )
            if "payment_address" not in columns:
                con.execute(
                    "ALTER TABLE invoices ADD COLUMN payment_address TEXT NOT NULL DEFAULT ''"
                )
            single_number_unique_index = False
            for index in con.execute("PRAGMA index_list(invoices)"):
                if index[2]:
                    indexed_columns = [
                        column[2]
                        for column in con.execute(
                            f"PRAGMA index_info('{index[1]}')"
                        )
                    ]
                    if indexed_columns == ["invoice_no"]:
                        single_number_unique_index = True
                        break
            if single_number_unique_index:
                con.execute(
                    """
                    CREATE TABLE invoices_migrated (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        invoice_no INTEGER NOT NULL,
                        invoice_date TEXT NOT NULL,
                        invoice_type TEXT NOT NULL,
                        customer_name TEXT,
                        grand_total REAL,
                        xlsx_path TEXT,
                        pdf_path TEXT,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        document_kind TEXT NOT NULL DEFAULT 'INVOICE',
                        qr_scanner_path TEXT,
                        payment_address TEXT NOT NULL DEFAULT '',
                        UNIQUE(document_kind, invoice_no)
                    )
                    """
                )
                con.execute(
                    """
                    INSERT INTO invoices_migrated (
                        id, invoice_no, invoice_date, invoice_type, customer_name,
                        grand_total, xlsx_path, pdf_path, created_at, document_kind,
                        qr_scanner_path, payment_address
                    )
                    SELECT id, invoice_no, invoice_date, invoice_type, customer_name,
                           grand_total, xlsx_path, pdf_path, created_at, document_kind,
                           qr_scanner_path, payment_address
                    FROM invoices
                    """
                )
                con.execute("DROP TABLE invoices")
                con.execute("ALTER TABLE invoices_migrated RENAME TO invoices")
            con.execute(
                """
                CREATE TABLE IF NOT EXISTS bank_profiles (
                    name TEXT PRIMARY KEY,
                    account TEXT NOT NULL DEFAULT '',
                    ifsc TEXT NOT NULL DEFAULT '',
                    payment_address TEXT NOT NULL DEFAULT '',
                    scanner_path TEXT
                )
                """
            )
            bank_columns = {
                row[1] for row in con.execute("PRAGMA table_info(bank_profiles)")
            }
            if "scanner_path" not in bank_columns:
                con.execute("ALTER TABLE bank_profiles ADD COLUMN scanner_path TEXT")
            con.execute(
                """
                CREATE TABLE IF NOT EXISTS qr_scanners (
                    id TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    image_path TEXT NOT NULL
                )
                """
            )
            con.execute(
                """
                CREATE TABLE IF NOT EXISTS seller_profile (
                    id INTEGER PRIMARY KEY CHECK (id = 1),
                    seller_name TEXT NOT NULL,
                    seller_business TEXT NOT NULL,
                    address_1 TEXT NOT NULL,
                    address_2 TEXT NOT NULL,
                    gstin TEXT NOT NULL,
                    state TEXT NOT NULL,
                    website_phone TEXT NOT NULL,
                    email_phone TEXT NOT NULL,
                    signatory TEXT NOT NULL,
                    jurisdiction TEXT NOT NULL
                )
                """
            )
            con.execute(
                """
                CREATE TABLE IF NOT EXISTS buyer_profiles (
                    name TEXT PRIMARY KEY,
                    address_1 TEXT NOT NULL DEFAULT '',
                    address_2 TEXT NOT NULL DEFAULT '',
                    mobile TEXT NOT NULL DEFAULT '',
                    gstin TEXT NOT NULL DEFAULT ''
                )
                """
            )
            con.execute(
                """
                CREATE TABLE IF NOT EXISTS document_drafts (
                    document_kind TEXT PRIMARY KEY,
                    state_json TEXT NOT NULL
                )
                """
            )
            con.execute(
                """
                CREATE TABLE IF NOT EXISTS invoice_payments (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    document_kind TEXT NOT NULL DEFAULT 'INVOICE',
                    invoice_no INTEGER NOT NULL,
                    paid_on TEXT NOT NULL,
                    amount REAL NOT NULL CHECK (amount > 0),
                    method TEXT NOT NULL,
                    reference TEXT NOT NULL DEFAULT '',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
            con.execute(
                """
                CREATE INDEX IF NOT EXISTS invoice_payments_document
                ON invoice_payments (document_kind, invoice_no)
                """
            )
            con.execute(
                """
                CREATE TABLE IF NOT EXISTS invoice_requests (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    invoice_no INTEGER NOT NULL,
                    requested_at TEXT NOT NULL,
                    request_percent REAL NOT NULL,
                    request_amount REAL NOT NULL,
                    pdf_path TEXT NOT NULL,
                    qr_path TEXT NOT NULL DEFAULT '',
                    payment_update_path TEXT NOT NULL DEFAULT '',
                    final_pdf_path TEXT NOT NULL DEFAULT ''
                )
                """
            )
            request_columns = {
                row[1] for row in con.execute("PRAGMA table_info(invoice_requests)")
            }
            if "payment_update_path" not in request_columns:
                con.execute(
                    "ALTER TABLE invoice_requests ADD COLUMN "
                    "payment_update_path TEXT NOT NULL DEFAULT ''"
                )
            if "final_pdf_path" not in request_columns:
                con.execute(
                    "ALTER TABLE invoice_requests ADD COLUMN "
                    "final_pdf_path TEXT NOT NULL DEFAULT ''"
                )
            con.execute(
                """
                CREATE INDEX IF NOT EXISTS invoice_requests_document
                ON invoice_requests (invoice_no, requested_at)
                """
            )


    def prune_missing_invoice_records(self):

        with sqlite3.connect(DB_PATH) as con:
            records = con.execute(
                "SELECT id, invoice_no, document_kind, xlsx_path, pdf_path FROM invoices"
            ).fetchall()
            saved_request_numbers = {
                invoice_no
                for invoice_no, pdf_path in con.execute(
                    "SELECT invoice_no, pdf_path FROM invoice_requests"
                )
                if pdf_path and Path(pdf_path).is_file()
            }
            missing_invoice_numbers = [
                (record_id, invoice_no, document_kind)
                for record_id, invoice_no, document_kind, xlsx_path, pdf_path in records
                if not any(
                    path and Path(path).is_file()
                    for path in (xlsx_path, pdf_path)
                )
                and not (
                    document_kind == "INVOICE"
                    and invoice_no in saved_request_numbers
                )
            ]
            con.executemany(
                "DELETE FROM invoice_payments WHERE document_kind = ? AND invoice_no = ?",
                ((kind, invoice_no) for _, invoice_no, kind in missing_invoice_numbers),
            )
            con.executemany(
                "DELETE FROM invoice_requests WHERE invoice_no = ?",
                (
                    (invoice_no,)
                    for _, invoice_no, kind in missing_invoice_numbers
                    if kind == "INVOICE"
                ),
            )
            con.executemany(
                "DELETE FROM invoices WHERE id = ?",
                ((record_id,) for record_id, _, _ in missing_invoice_numbers),
            )


    def next_invoice_number(self, document_kind=None):

        self.prune_missing_invoice_records()
        if document_kind is None:
            document_kind = document_kind_for_selection(
                self.document_kind.currentIndex()
            )

        with sqlite3.connect(DB_PATH) as con:

            row = con.execute(
                "SELECT MAX(invoice_no) FROM invoices WHERE document_kind = ?",
                (document_kind,),
            ).fetchone()

        return max(START_INVOICE_NUMBER, (row[0] or 0) + 1)


    def saved_invoice_file(self, invoice_no, document_kind=None):

        if document_kind is None:
            document_kind = document_kind_for_selection(
                self.document_kind.currentIndex()
            )

        with sqlite3.connect(DB_PATH) as con:
            row = con.execute(
                """SELECT xlsx_path, customer_name FROM invoices
                   WHERE invoice_no = ? AND document_kind = ?""",
                (invoice_no, document_kind),
            ).fetchone()

        if not row or not row[0]:
            return None, None

        return Path(row[0]), row[1]


    def seller_profile_fields(self):

        return (
            self.seller_name,
            self.seller_business,
            self.seller_address_1,
            self.seller_address_2,
            self.seller_gstin,
            self.seller_state,
            self.seller_website_phone,
            self.seller_email_phone,
            self.seller_signatory,
            self.jurisdiction,
        )


    def load_seller_profile(self):

        with sqlite3.connect(DB_PATH) as con:
            profile = con.execute(
                """
                SELECT seller_name, seller_business, address_1, address_2,
                       gstin, state, website_phone, email_phone, signatory,
                       jurisdiction
                FROM seller_profile WHERE id = 1
                """
            ).fetchone()
        if not profile:
            return

        for field, value in zip(self.seller_profile_fields(), profile):
            field.setText(value)


    def save_seller_profile(self):

        with sqlite3.connect(DB_PATH) as con:
            con.execute(
                """
                INSERT INTO seller_profile (
                    id, seller_name, seller_business, address_1, address_2,
                    gstin, state, website_phone, email_phone, signatory,
                    jurisdiction
                ) VALUES (1, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(id) DO UPDATE SET
                    seller_name = excluded.seller_name,
                    seller_business = excluded.seller_business,
                    address_1 = excluded.address_1,
                    address_2 = excluded.address_2,
                    gstin = excluded.gstin,
                    state = excluded.state,
                    website_phone = excluded.website_phone,
                    email_phone = excluded.email_phone,
                    signatory = excluded.signatory,
                    jurisdiction = excluded.jurisdiction
                """,
                tuple(field.text().strip() for field in self.seller_profile_fields()),
            )


    def refresh_buyer_profiles(self, selected_name=""):

        with sqlite3.connect(DB_PATH) as con:
            profiles = con.execute(
                "SELECT name FROM buyer_profiles ORDER BY name COLLATE NOCASE"
            ).fetchall()

        self.buyer_selector.blockSignals(True)
        self.buyer_selector.clear()
        self.buyer_selector.addItem("Select saved buyer", "")
        for (name,) in profiles:
            self.buyer_selector.addItem(name, name)
        selected_index = self.buyer_selector.findData(selected_name)
        self.buyer_selector.setCurrentIndex(max(0, selected_index))
        self.buyer_selector.blockSignals(False)


    def load_buyer_profile(self, index):

        profile_name = self.buyer_selector.itemData(index)
        if not profile_name:
            return

        with sqlite3.connect(DB_PATH) as con:
            profile = con.execute(
                """
                SELECT name, address_1, address_2, mobile, gstin
                FROM buyer_profiles WHERE name = ?
                """,
                (profile_name,),
            ).fetchone()
        if not profile:
            return

        for field, value in zip(
            (
                self.customer_name,
                self.customer_address,
                self.customer_address_2,
                self.customer_mobile,
                self.customer_gstin,
            ),
            profile,
        ):
            field.setText(value)


    def save_buyer_profile(self):

        name = self.customer_name.text().strip()
        if not name:
            QMessageBox.warning(self, "Buyer Name Required", "Enter a buyer name first.")
            return
        mobile = self.customer_mobile.text().strip()
        if not normalize_indian_mobile(mobile):
            QMessageBox.warning(
                self,
                "Invalid Buyer Mobile",
                "Enter a valid Indian mobile number with 10 digits, optionally prefixed by +91.",
            )
            self.customer_mobile.setFocus()
            return

        with sqlite3.connect(DB_PATH) as con:
            con.execute(
                """
                INSERT INTO buyer_profiles (name, address_1, address_2, mobile, gstin)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(name) DO UPDATE SET
                    address_1 = excluded.address_1,
                    address_2 = excluded.address_2,
                    mobile = excluded.mobile,
                    gstin = excluded.gstin
                """,
                (
                    name,
                    self.customer_address.text().strip(),
                    self.customer_address_2.text().strip(),
                    mobile,
                    self.customer_gstin.text().strip(),
                ),
            )
        self.refresh_buyer_profiles(name)
        self.statusBar().showMessage(f"Buyer details saved: {name}")


    def delete_buyer_profile(self):

        profile_name = self.buyer_selector.currentData()
        if not profile_name:
            QMessageBox.information(
                self, "Select a Saved Buyer", "Choose a saved buyer to delete."
            )
            return

        answer = QMessageBox.question(
            self,
            "Delete Buyer Details",
            f"Delete the saved buyer '{profile_name}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return

        with sqlite3.connect(DB_PATH) as con:
            con.execute("DELETE FROM buyer_profiles WHERE name = ?", (profile_name,))
        self.refresh_buyer_profiles()


    def new_buyer_profile(self):

        self.buyer_selector.setCurrentIndex(0)
        for field in (
            self.customer_name,
            self.customer_address,
            self.customer_address_2,
            self.customer_mobile,
            self.customer_gstin,
        ):
            field.clear()


    def refresh_bank_profiles(self, selected_name=None):

        if selected_name is None:
            selected_name = self.bank_name.text().strip()
        with sqlite3.connect(DB_PATH) as con:
            profiles = con.execute(
                "SELECT name FROM bank_profiles ORDER BY name COLLATE NOCASE"
            ).fetchall()

        self.bank_selector.blockSignals(True)
        self.bank_selector.clear()
        self.bank_selector.addItem("Manual / not saved", "")
        for (name,) in profiles:
            self.bank_selector.addItem(name, name)
        selected_index = self.bank_selector.findData(selected_name)
        self.bank_selector.setCurrentIndex(max(0, selected_index))
        self.bank_selector.blockSignals(False)
        if selected_index > 0:
            self.load_bank_profile(selected_index)


    def load_bank_profile(self, index):

        profile_name = self.bank_selector.itemData(index)
        if not profile_name:
            return

        with sqlite3.connect(DB_PATH) as con:
            profile = con.execute(
                """
                SELECT name, account, ifsc, payment_address
                FROM bank_profiles WHERE name = ?
                """,
                (profile_name,),
            ).fetchone()
        if not profile:
            return

        self.bank_name.setText(profile[0])
        self.bank_account.setText(profile[1])
        self.bank_ifsc.setText(profile[2])
        self.bank_upi_id.setText(profile[3])


    def save_bank_profile(self):

        name = self.bank_name.text().strip()
        if not name:
            QMessageBox.warning(self, "Bank Name Required", "Enter a bank name first.")
            return

        with sqlite3.connect(DB_PATH) as con:
            con.execute(
                """
                INSERT INTO bank_profiles (name, account, ifsc, payment_address)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(name) DO UPDATE SET
                    account = excluded.account,
                    ifsc = excluded.ifsc,
                    payment_address = excluded.payment_address
                """,
                (
                    name,
                    self.bank_account.text().strip(),
                    self.bank_ifsc.text().strip(),
                    self.bank_upi_id.text().strip(),
                ),
            )
        self.refresh_bank_profiles(name)
        self.statusBar().showMessage(f"Bank profile saved: {name}")


    def new_bank_profile(self):

        self.bank_selector.setCurrentIndex(0)
        self.bank_name.clear()
        self.bank_account.clear()
        self.bank_ifsc.clear()
        self.bank_upi_id.clear()


    def delete_bank_profile(self):

        profile_name = self.bank_selector.currentData()
        if not profile_name:
            QMessageBox.information(
                self, "Select a Saved Bank", "Choose a saved bank profile to delete."
            )
            return

        answer = QMessageBox.question(
            self,
            "Delete Bank Profile",
            f"Delete the saved bank profile '{profile_name}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return

        with sqlite3.connect(DB_PATH) as con:
            con.execute("DELETE FROM bank_profiles WHERE name = ?", (profile_name,))
        self.refresh_bank_profiles("")
        self.new_bank_profile()


    def select_bank_profile(self, name):

        profile_index = self.bank_selector.findData(name)
        self.bank_selector.setCurrentIndex(max(0, profile_index))


    def payment_uri(self, invoice_no, amount=None):

        uri = (
            "upi://pay?"
            f"pa={quote(self.bank_upi_id.text().strip())}"
            f"&pn={quote(self.seller_name.text())}"
            "&cu=INR"
            f"&tn={quote(str(invoice_no))}"
        )
        if amount is not None:
            uri += f"&am={money(amount):.2f}"
        return uri


    def invoice_storage_root(self, document_kind, invoice_no, grand_total=None):

        root = (
            self.custom_output_dir
            or (
                PRIVATE_INVOICE_DIR
                if self.private_invoice_checkbox.isChecked()
                else INVOICE_DIR
            )
        )
        if document_kind != "INVOICE":
            return root

        if grand_total is None:
            with sqlite3.connect(DB_PATH) as con:
                invoice = con.execute(
                    "SELECT grand_total FROM invoices "
                    "WHERE invoice_no = ? AND document_kind = 'INVOICE'",
                    (invoice_no,),
                ).fetchone()
            grand_total = invoice[0] if invoice else self.calculate()[3]

        if money(grand_total or 0) > self.invoice_paid_total(invoice_no):
            return root / "Pending"
        return root


    def quotation_payment_qr_path(self, invoice_no, invoice_date, grand_total=None):

        document_kind = document_kind_for_selection(
            self.document_kind.currentIndex()
        )
        document_root = self.invoice_storage_root(
            document_kind, invoice_no, grand_total
        )
        return (
            document_root
            / document_folder_for_selection(self.document_kind.currentIndex())
            / "QR"
            / str(invoice_date.year)
            / invoice_filename(
                invoice_no,
                self.customer_name.text(),
                "_payment.png",
                document_kind_for_selection(
                    self.document_kind.currentIndex()
                ),
            )
        )


    def load_invoice(self):

        try:
            invoice_no = int(self.invoice_no.text())
        except ValueError:
            return

        with sqlite3.connect(DB_PATH) as con:
            row = con.execute(
                """
                SELECT invoice_date, invoice_type, xlsx_path, grand_total,
                        document_kind, qr_scanner_path, payment_address
                FROM invoices
                WHERE invoice_no = ? AND document_kind = ?
                """,
                (
                    invoice_no,
                    document_kind_for_selection(self.document_kind.currentIndex()),
                ),
            ).fetchone()

        if not row:
            self.loaded_invoice_no = None
            self.refresh_invoice_payments()
            self.refresh_invoice_request_history()
            self.statusBar().showMessage(
                f"No saved invoice found for number {invoice_no}"
            )
            return

        workbook_path = Path(row[2])
        if not workbook_path.is_file():
            QMessageBox.warning(
                self,
                "Invoice File Missing",
                f"The saved workbook could not be found:\n{workbook_path}",
            )
            return
        self.reset_receipt_amount_autofill()
        self.private_invoice_checkbox.setChecked(
            workbook_path.is_relative_to(PRIVATE_INVOICE_DIR)
        )

        sheet = load_workbook(workbook_path, data_only=True).active

        def cell_text(address, prefix=""):
            value = sheet[address].value
            text = "" if value is None else str(value)
            if prefix and text.startswith(prefix):
                return text[len(prefix):].strip()
            return text

        self.loaded_invoice_no = invoice_no
        self.document_kind.setCurrentIndex(
            document_kind_selection_for_kind(row[4])
        )
        self.buyer_selector.setCurrentIndex(0)
        if row[4] in ("INVOICE", "QUOTATION"):
            payment_address = (
                row[6]
                or cell_text("G49", "UPI ID:")
                or cell_text("G3", "UPI ID:")
            )
            if not payment_address:
                payment_address = cell_text("A54", "Payee address:")
            if payment_address.startswith("UPI ID:"):
                payment_address = payment_address.removeprefix("UPI ID:").strip()
            elif payment_address.startswith("Payee address:"):
                payment_address = payment_address.removeprefix("Payee address:").strip()
            if payment_address == "Uploaded scanner selected":
                payment_address = ""
            self.bank_upi_id.setText(payment_address)
            payment_match = re.search(
                r"(?:Advance|Scan to pay|Collect|Balance|Payment Request)\s+([\d.]+)%",
                cell_text("H44") or cell_text("G48") or cell_text("E2") or cell_text("A53"),
            )
            self.payment_percent.setValue(
                float(payment_match.group(1)) if payment_match else 50
            )
        self.invoice_date.setDate(QDate.fromString(row[0], "yyyy-MM-dd"))
        self.invoice_type.setCurrentIndex(0 if row[1] == "IGST" else 1)
        website_contact = "   ".join(
            str(value).strip()
            for value in (sheet["A8"].value, sheet["E8"].value)
            if value is not None and str(value).strip()
        )
        email_contact = "   ".join(
            str(value).strip()
            for value in (sheet["A9"].value, sheet["E9"].value)
            if value is not None and str(value).strip()
        )
        seller_values = (
            (self.seller_name, sheet["A2"].value),
            (self.seller_business, sheet["A3"].value),
            (self.seller_address_1, sheet["A4"].value),
            (self.seller_address_2, sheet["A5"].value),
            (self.seller_gstin, cell_text("A6", "GSTIN :") or cell_text("A6", "GSTIN:")),
            (self.seller_state, sheet["A7"].value),
            (self.seller_website_phone, website_contact),
            (self.seller_email_phone, email_contact),
        )
        for field, value in seller_values:
            if value is not None and str(value).strip():
                field.setText(str(value))

        self.customer_name.setText(cell_text("A11"))
        customer_address = sheet["A12"].value
        if not customer_address:
            legacy_address = sheet["A14"].value
            if legacy_address and not str(legacy_address).startswith("Mobile:"):
                customer_address = legacy_address
        address_lines = str(customer_address or "").splitlines()
        self.customer_address.setText(address_lines[0] if address_lines else "")
        self.customer_address_2.setText("\n".join(address_lines[1:]))
        self.po_number.setText(cell_text("G14") or cell_text("F13"))
        mobile = cell_text("A14", "Mobile:")
        gstin = cell_text("A15", "Gst no:")
        legacy_contact = cell_text("A13")
        if legacy_contact:
            mobile_match = re.search(r"Mobile\s*:\s*(.*)$", legacy_contact, re.IGNORECASE)
            gstin_match = re.search(
                r"GSTIN\s*:\s*(.*?)(?:\s+Mobile\s*:|$)",
                legacy_contact,
                re.IGNORECASE,
            )
            if mobile_match:
                mobile = mobile or mobile_match.group(1).strip()
            if gstin_match:
                gstin = gstin or gstin_match.group(1).strip()
        self.customer_mobile.setText(mobile)
        self.customer_gstin.setText(gstin)
        if row[4] in ("INVOICE", "QUOTATION"):
            for field, address, prefix in (
                (self.bank_name, "D44", "Bank Name:"),
                (self.bank_account, "D45", "A/C No:"),
                (self.bank_ifsc, "D46", "IFSC:"),
            ):
                value = cell_text(address, prefix)
                if value:
                    field.setText(value)
            self.seller_signatory.setText(
                cell_text("D47") or cell_text("F47") or self.seller_signatory.text()
            )
        for field, address in ((self.bundle_count, "A44"), (self.through, "A45")):
            value = sheet[address].value
            if value:
                field.setText(str(value).split(":", 1)[-1].strip())
        for field, address in (
            (self.bank_name, "F44"),
            (self.bank_account, "F45"),
            (self.bank_ifsc, "F46"),
            (self.seller_signatory, "F47"),
            (self.jurisdiction, "A51"),
        ):
            value = sheet[address].value
            if value is not None and str(value).strip():
                field.setText(str(value).strip())
        self.select_bank_profile(self.bank_name.text())
        if row[6]:
            self.bank_upi_id.setText(row[6])

        declaration_parts = [
            str(sheet[address].value).strip()
            for address in ("A48", "A49")
            if sheet[address].value and str(sheet[address].value).strip()
        ]
        if declaration_parts:
            self.declaration.setPlainText(" ".join(declaration_parts))

        tax_row = next(
            (
                candidate
                for candidate in (38, 37)
                if isinstance(sheet[f"D{candidate}"].value, (int, float))
            ),
            None,
        )
        rate_cell = sheet[f"D{tax_row}"].value if tax_row else None
        if rate_cell is None:
            tax_label = str(sheet["B29"].value or "")
            tax_match = re.search(r"(\d+(?:\.\d+)?)\s*%", tax_label)
            subtotal_value = sheet["H28"].value or 0
            if tax_match:
                gst_rate = float(tax_match.group(1))
            elif subtotal_value:
                multiplier = 100 if row[1] == "IGST" else 200
                gst_rate = max(0, (float(row[3]) / float(subtotal_value) - 1) * multiplier)
            else:
                gst_rate = self.gst_rate.value()
        else:
            gst_rate = float(rate_cell) * (100 if row[1] == "IGST" else 200)
        self.gst_rate.setValue(gst_rate)
        hsn_code = (sheet[f"A{tax_row}"].value if tax_row else None) or sheet["A38"].value or sheet["A37"].value
        if hsn_code:
            self.hsn_code.setText(str(hsn_code))

        self.items.setRowCount(0)
        for source_row in range(17, 28):
            description = sheet.cell(source_row, 2).value
            if not description:
                continue

            self.add_item_row()
            target_row = self.items.rowCount() - 1
            self.items.blockSignals(True)
            for column, source_column in ((1, 2), (2, 4), (3, 5), (4, 6), (5, 7)):
                value = sheet.cell(source_row, source_column).value
                self.items.item(target_row, column).setText(
                    "PCS" if column == 5 and not str(value or "").strip()
                    else "" if value is None else str(value)
                )
            self.items.blockSignals(False)

        self.refresh_item_row_controls()
        self.refresh()
        self.refresh_invoice_payments()
        self.refresh_invoice_request_history()
        self.statusBar().showMessage(f"Invoice {invoice_no} loaded for editing")


    # ========================================================
    # UI
    # ========================================================

    def sidebar_action_icon(self, kind):
        pixmap = QPixmap(24, 24)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        if kind == "plus":
            painter.setPen(QPen(QColor("#21866b"), 2.4))
            painter.drawLine(12, 5, 12, 19)
            painter.drawLine(5, 12, 19, 12)
        else:
            painter.setPen(QPen(QColor("#36515d"), 1.7))
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.drawRoundedRect(QRectF(3.5, 2, 15, 20), 2, 2)

        if kind == "plus":
            pass
        elif kind == "create":
            painter.drawLine(7, 8, 14, 8)
            painter.drawLine(7, 11, 13, 11)
            painter.setPen(QPen(QColor("#21866b"), 2.2))
            painter.drawLine(18, 15, 18, 22)
            painter.drawLine(14.5, 18.5, 21.5, 18.5)
        else:
            painter.setPen(QColor("#21866b"))
            painter.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
            painter.drawText(
                QRectF(5, 2, 13, 13),
                Qt.AlignmentFlag.AlignCenter,
                '"',
            )
            painter.setPen(QPen(QColor("#36515d"), 1.5))
            painter.drawLine(7, 15, 15, 15)
            painter.drawLine(7, 18, 15, 18)

        painter.end()
        return QIcon(pixmap)


    def create_ui(self):
        central = QWidget()
        root_layout = QHBoxLayout(central)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        sidebar = QWidget()
        sidebar.setFixedWidth(196)
        sidebar.setStyleSheet("background: #e9eff2;")
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(14, 20, 14, 16)
        sidebar_layout.setSpacing(10)
        brand = QLabel("ADOODS\nINVOICE MANAGER")
        brand.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        brand.setStyleSheet("color: #193f4c; padding: 8px 6px 18px;")
        sidebar_layout.addWidget(brand)

        self.sidebar_navigation = QListWidget()
        navigation_items = (
            ("Create", self.sidebar_action_icon("plus")),
            ("invoices", self.sidebar_action_icon("create")),
            ("Quotations", self.sidebar_action_icon("quotation")),
            ("Finance", self.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon)),
        )
        for label, icon in navigation_items:
            item = QListWidgetItem(icon, label)
            self.sidebar_navigation.addItem(item)
        self.sidebar_navigation.setSpacing(4)
        self.sidebar_navigation.setIconSize(QSize(18, 18))
        self.sidebar_navigation.setCurrentRow(0)
        self.sidebar_navigation.setStyleSheet(
            "QListWidget { border: 0; background: transparent; padding: 2px; }"
            "QListWidget::item { padding: 12px 10px; border-radius: 4px; "
            "color: #36515d; font-weight: 600; }"
            "QListWidget::item:selected { background: #d0e6dd; color: #155b46; }"
            "QListWidget::item:hover { background: #dce7e8; }"
        )
        sidebar_layout.addWidget(self.sidebar_navigation, 1)
        sidebar_note = QLabel("GST invoices\nReceipts & reporting")
        sidebar_note.setStyleSheet("color: #6f8188; padding: 8px 6px;")
        sidebar_layout.addWidget(sidebar_note)
        root_layout.addWidget(sidebar)

        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(18, 16, 18, 16)
        content_layout.setSpacing(8)
        title = QLabel("Invoice Manager")
        title.setFont(QFont("Segoe UI", 20, QFont.Weight.DemiBold))
        content_layout.addWidget(title)
        subtitle = QLabel("Create and manage GST invoices")
        subtitle.setStyleSheet("color:#667783;")
        content_layout.addWidget(subtitle)

        self.main_tabs = QStackedWidget()
        self.invoice_tab = QWidget()
        self.create_invoice_tab()
        self.main_tabs.addWidget(self.invoice_tab)
        self.create_pending_invoices_tab()
        self.main_tabs.addWidget(self.pending_invoices_tab)
        self.create_quotation_register_tab()
        self.main_tabs.addWidget(self.quotation_register_tab)
        self.create_finance_tab()
        self.main_tabs.addWidget(self.finance_tab)
        self.sidebar_navigation.currentRowChanged.connect(
            self.main_tabs.setCurrentIndex
        )
        self.main_tabs.currentChanged.connect(self.sync_sidebar_navigation)
        self.main_tabs.currentChanged.connect(self.refresh_finance_if_visible)
        content_layout.addWidget(self.main_tabs, 1)
        root_layout.addWidget(content, 1)
        self.setCentralWidget(central)
        return


    def finance_metric_card(self, title):

        card = QFrame()
        card.setStyleSheet(
            "QFrame { background: #ffffff; border: 1px solid #d9e2e6; "
            "border-radius: 5px; }"
        )
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(15, 12, 15, 12)
        card_layout.setSpacing(5)
        title_label = QLabel(title.upper())
        title_label.setStyleSheet(
            "color: #657780; border: 0; font-size: 8pt; font-weight: 600;"
        )
        value_label = QLabel("₹0.00")
        value_label.setFont(QFont("Segoe UI", 17, QFont.Weight.DemiBold))
        value_label.setStyleSheet("color: #193f4c; border: 0;")
        detail_label = QLabel("")
        detail_label.setStyleSheet("color: #72828a; border: 0; font-size: 9pt;")
        card_layout.addWidget(title_label)
        card_layout.addWidget(value_label)
        card_layout.addWidget(detail_label)
        return card, value_label, detail_label


    def create_finance_tab(self):

        self.finance_tab = QWidget()
        tab_layout = QVBoxLayout(self.finance_tab)
        tab_layout.setContentsMargins(0, 0, 0, 0)
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        finance_content = QWidget()
        scroll_area.setWidget(finance_content)
        tab_layout.addWidget(scroll_area)

        layout = QVBoxLayout(finance_content)
        layout.setContentsMargins(18, 14, 18, 16)
        layout.setSpacing(12)

        header = QHBoxLayout()
        heading = QLabel("Financial overview")
        heading.setFont(QFont("Segoe UI", 17, QFont.Weight.DemiBold))
        header.addWidget(heading)
        header.addStretch()
        header.addWidget(QLabel("Calendar year"))
        self.finance_year_selector = NoWheelComboBox()
        self.finance_year_selector.setMinimumWidth(110)
        header.addWidget(self.finance_year_selector)
        refresh_button = QPushButton("Refresh")
        refresh_button.clicked.connect(self.refresh_finance_dashboard)
        header.addWidget(refresh_button)
        layout.addLayout(header)

        metrics = QHBoxLayout()
        metrics.setSpacing(10)
        self.finance_collected_card, self.finance_collected_value, self.finance_collected_detail = (
            self.finance_metric_card("Collected this year")
        )
        self.finance_sales_card, self.finance_sales_value, self.finance_sales_detail = (
            self.finance_metric_card("Invoiced this year")
        )
        self.finance_outstanding_card, self.finance_outstanding_value, self.finance_outstanding_detail = (
            self.finance_metric_card("Outstanding now")
        )
        self.finance_growth_card, self.finance_growth_value, self.finance_growth_detail = (
            self.finance_metric_card("Collection growth")
        )
        for card in (
            self.finance_collected_card,
            self.finance_sales_card,
            self.finance_outstanding_card,
            self.finance_growth_card,
        ):
            metrics.addWidget(card, 1)
        layout.addLayout(metrics)

        charts = QHBoxLayout()
        trend_box = QGroupBox("Monthly receipts")
        trend_layout = QVBoxLayout(trend_box)
        self.finance_monthly_chart = FinanceMonthlyChart()
        trend_layout.addWidget(self.finance_monthly_chart)
        self.finance_monthly_insight = QLabel("Receipts recorded in the selected year")
        self.finance_monthly_insight.setStyleSheet("color: #657780;")
        trend_layout.addWidget(self.finance_monthly_insight)
        charts.addWidget(trend_box, 2)

        mix_box = QGroupBox("Invoice position")
        mix_layout = QVBoxLayout(mix_box)
        self.finance_donut_chart = FinanceDonutChart()
        mix_layout.addWidget(self.finance_donut_chart, 1)
        self.finance_mix_collected = QLabel("Collected: ₹0.00")
        self.finance_mix_pending = QLabel("Outstanding: ₹0.00")
        self.finance_mix_collected.setStyleSheet("color: #21866b; font-weight: 600;")
        self.finance_mix_pending.setStyleSheet("color: #b7792f; font-weight: 600;")
        mix_layout.addWidget(self.finance_mix_collected)
        mix_layout.addWidget(self.finance_mix_pending)
        charts.addWidget(mix_box, 1)
        layout.addLayout(charts, 1)

        detail_row = QHBoxLayout()
        debtors_box = QGroupBox("Largest outstanding customer balances")
        debtors_layout = QVBoxLayout(debtors_box)
        self.finance_debtors_table = QTableWidget(0, 3)
        self.finance_debtors_table.setHorizontalHeaderLabels(
            ["Customer", "Open invoices", "Balance due"]
        )
        self.finance_debtors_table.setEditTriggers(
            QTableWidget.EditTrigger.NoEditTriggers
        )
        self.finance_debtors_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )
        self.finance_debtors_table.verticalHeader().setVisible(False)
        self.finance_debtors_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.Stretch
        )
        self.finance_debtors_table.setMinimumHeight(170)
        debtors_layout.addWidget(self.finance_debtors_table)
        detail_row.addWidget(debtors_box, 3)

        insights_box = QGroupBox("At a glance")
        insights_layout = QVBoxLayout(insights_box)
        self.finance_overdue_label = QLabel("Overdue invoices: 0")
        self.finance_overdue_label.setWordWrap(True)
        self.finance_overdue_label.setStyleSheet("font-weight: 600; color: #193f4c;")
        self.finance_year_insight = QLabel("Year-on-year comparison will appear here.")
        self.finance_year_insight.setWordWrap(True)
        self.finance_year_insight.setStyleSheet("color: #657780;")
        self.finance_top_month_insight = QLabel("Best collection month: --")
        self.finance_top_month_insight.setWordWrap(True)
        insights_layout.addWidget(self.finance_overdue_label)
        insights_layout.addWidget(self.finance_year_insight)
        insights_layout.addWidget(self.finance_top_month_insight)
        insights_layout.addStretch()
        detail_row.addWidget(insights_box, 2)
        layout.addLayout(detail_row)

        self.finance_year_selector.currentTextChanged.connect(
            self.refresh_finance_dashboard
        )


    def sync_sidebar_navigation(self, index):

        if self.sidebar_navigation.currentRow() != index:
            self.sidebar_navigation.setCurrentRow(index)


    def refresh_finance_if_visible(self, index):

        if index == 3:
            self.refresh_finance_dashboard()


    def refresh_finance_dashboard(self, selected_year=None):

        if not hasattr(self, "finance_year_selector"):
            return
        if selected_year is None:
            selected_year = self.finance_year_selector.currentText()

        with sqlite3.connect(DB_PATH) as con:
            invoice_rows = con.execute(
                """
                SELECT i.invoice_no, i.invoice_date, COALESCE(i.customer_name, ''),
                       COALESCE(i.grand_total, 0), COALESCE(SUM(p.amount), 0)
                FROM invoices AS i
                LEFT JOIN invoice_payments AS p
                  ON p.document_kind = i.document_kind
                 AND p.invoice_no = i.invoice_no
                WHERE i.document_kind = 'INVOICE'
                GROUP BY i.id
                ORDER BY i.invoice_date DESC
                """
            ).fetchall()
            payment_rows = con.execute(
                """
                SELECT p.paid_on, p.amount
                FROM invoice_payments AS p
                INNER JOIN invoices AS i
                  ON i.document_kind = p.document_kind
                 AND i.invoice_no = p.invoice_no
                WHERE p.document_kind = 'INVOICE'
                """
            ).fetchall()

        available_years = {datetime.now().year}
        available_years.update(
            int(invoice_date[:4])
            for _, invoice_date, _, _, _ in invoice_rows
            if invoice_date and invoice_date[:4].isdigit()
        )
        available_years.update(
            int(paid_on[:4])
            for paid_on, _ in payment_rows
            if paid_on and paid_on[:4].isdigit()
        )
        year_text = str(selected_year or datetime.now().year)
        self.finance_year_selector.blockSignals(True)
        self.finance_year_selector.clear()
        self.finance_year_selector.addItems(
            [str(year) for year in sorted(available_years, reverse=True)]
        )
        if self.finance_year_selector.findText(year_text) >= 0:
            self.finance_year_selector.setCurrentText(year_text)
        self.finance_year_selector.blockSignals(False)
        year_text = self.finance_year_selector.currentText()
        year = int(year_text) if year_text.isdigit() else datetime.now().year
        today = datetime.now().date()
        comparison_cutoff = (
            f"{year - 1}-{today:%m-%d}"
            if year == today.year else f"{year - 1}-12-31"
        )
        comparison_period = (
            f"{year - 1} through {today:%b %d}"
            if year == today.year else str(year - 1)
        )

        yearly_sales = sum(
            float(total or 0)
            for _, invoice_date, _, total, _ in invoice_rows
            if invoice_date and invoice_date.startswith(year_text)
        )
        monthly_receipts = [0.0] * 12
        yearly_receipts = 0.0
        previous_year_receipts = 0.0
        for paid_on, amount in payment_rows:
            if not paid_on or len(paid_on) < 7:
                continue
            payment_year = paid_on[:4]
            if payment_year == year_text:
                yearly_receipts += float(amount or 0)
                try:
                    month_index = int(paid_on[5:7]) - 1
                except ValueError:
                    continue
                if 0 <= month_index < 12:
                    monthly_receipts[month_index] += float(amount or 0)
            elif payment_year == str(year - 1) and paid_on <= comparison_cutoff:
                previous_year_receipts += float(amount or 0)

        total_collected = 0.0
        total_outstanding = 0.0
        overdue_count = 0
        pending_by_customer = {}
        for _, invoice_date, customer, total, paid in invoice_rows:
            invoice_total = max(0.0, float(total or 0))
            received = max(0.0, float(paid or 0))
            allocated = min(invoice_total, received)
            balance = max(0.0, invoice_total - received)
            total_collected += allocated
            total_outstanding += balance
            if balance > 0.005:
                buyer = customer.strip() or "Unnamed buyer"
                entry = pending_by_customer.setdefault(buyer, [0, 0.0])
                entry[0] += 1
                entry[1] += balance
                try:
                    invoice_day = datetime.strptime(invoice_date, "%Y-%m-%d").date()
                    if (today - invoice_day).days > 30:
                        overdue_count += 1
                except (TypeError, ValueError):
                    pass

        self.finance_collected_value.setText(f"₹{yearly_receipts:,.2f}")
        self.finance_collected_detail.setText(f"Cash received during {year}")
        self.finance_sales_value.setText(f"₹{yearly_sales:,.2f}")
        self.finance_sales_detail.setText(f"Invoices dated in {year}")
        self.finance_outstanding_value.setText(f"₹{total_outstanding:,.2f}")
        open_count = sum(count for count, _ in pending_by_customer.values())
        self.finance_outstanding_detail.setText(f"{open_count} invoice(s) still open")

        if previous_year_receipts > 0:
            growth = (yearly_receipts - previous_year_receipts) / previous_year_receipts * 100
            growth_prefix = "+" if growth >= 0 else ""
            self.finance_growth_value.setText(f"{growth_prefix}{growth:.1f}%")
            self.finance_growth_detail.setText(
                f"vs {comparison_period} (₹{previous_year_receipts:,.2f})"
            )
            self.finance_year_insight.setText(
                f"Receipts are {growth_prefix}{growth:.1f}% versus {comparison_period}."
            )
        else:
            self.finance_growth_value.setText("--")
            self.finance_growth_detail.setText(f"No {year - 1} receipt history")
            self.finance_year_insight.setText(
                f"No comparable receipt history for {year - 1}."
            )

        self.finance_monthly_chart.set_values(monthly_receipts)
        nonzero_months = [
            (amount, FinanceMonthlyChart.MONTHS[index])
            for index, amount in enumerate(monthly_receipts)
            if amount > 0
        ]
        if nonzero_months:
            best_amount, best_month = max(nonzero_months)
            self.finance_top_month_insight.setText(
                f"Best collection month: {best_month} · ₹{best_amount:,.2f}"
            )
            latest_index = max(index for index, amount in enumerate(monthly_receipts) if amount > 0)
            if latest_index > 0 and monthly_receipts[latest_index - 1] > 0:
                prior = monthly_receipts[latest_index - 1]
                change = (monthly_receipts[latest_index] - prior) / prior * 100
                change_prefix = "+" if change >= 0 else ""
                self.finance_monthly_insight.setText(
                    f"Latest active month is {change_prefix}{change:.1f}% versus the prior month."
                )
            else:
                self.finance_monthly_insight.setText(
                    f"Receipts recorded in {year}; strongest month was {best_month}."
                )
        else:
            self.finance_top_month_insight.setText("Best collection month: --")
            self.finance_monthly_insight.setText(f"No receipts recorded in {year}.")

        self.finance_donut_chart.set_values(total_collected, total_outstanding)
        self.finance_mix_collected.setText(f"Collected: ₹{total_collected:,.2f}")
        self.finance_mix_pending.setText(f"Outstanding: ₹{total_outstanding:,.2f}")
        self.finance_overdue_label.setText(
            f"Open more than 30 days: {overdue_count} invoice(s)"
        )
        sorted_customers = sorted(
            pending_by_customer.items(), key=lambda item: item[1][1], reverse=True
        )[:5]
        self.finance_debtors_table.setRowCount(len(sorted_customers))
        for row, (customer, (invoice_count, balance)) in enumerate(sorted_customers):
            for column, value in enumerate((
                customer,
                str(invoice_count),
                f"₹{balance:,.2f}",
            )):
                self.finance_debtors_table.setItem(
                    row, column, QTableWidgetItem(value)
                )

        return

        self.invoice_type = NoWheelComboBox()

        self.invoice_type.addItems([
            "IGST",
            "CGST + SGST"
        ])

        self.invoice_type.currentIndexChanged.connect(
            self.refresh
        )
        self.invoice_date = NoWheelDateEdit()

        self.invoice_date.setCalendarPopup(
            True
        )

        self.invoice_date.setDisplayFormat(
            "dd/MM/yyyy"
        )

        self.invoice_date.dateChanged.connect(
            self.refresh
        )

        self.invoice_no.textChanged.connect(
            self.refresh
        )


        grid.addWidget(
            QLabel("Invoice Template"),
            0, 0
        )

        grid.addWidget(
            self.invoice_type,
            0, 1
        )


        grid.addWidget(
            QLabel("Invoice No"),
            0, 2
        )

        grid.addWidget(
            self.invoice_no,
            0, 3
        )


        grid.addWidget(
            QLabel("Date"),
            1, 0
        )

        grid.addWidget(
            self.invoice_date,
            1, 1
        )


        layout.addWidget(
            invoice_box
        )


        seller_box = QGroupBox(
            "Seller Details"
        )

        seller_form = QFormLayout(
            seller_box
        )

        self.seller_name = QLineEdit(
            "ADOODS GARMENTS"
        )

        self.seller_address = QLineEdit(
            "88-1/50M, 1st FLOOR NEAR SRINIVASA THEATRE, "
            "ULAVARSANDHAI ROAD, HOSUR - 635 109"
        )


        self.seller_gstin = QLineEdit(
            "33DJRPP6135G1ZW"
        )


        seller_form.addRow(
            "Company Name",
            self.seller_name
        )

        seller_form.addRow(
            "Address",
            self.seller_address
        )

        seller_form.addRow(
            "GSTIN",
            self.seller_gstin
        )


        layout.addWidget(
            seller_box
        )


        # ----------------------------------------------------
        # CUSTOMER
        # ----------------------------------------------------

        customer_box = QGroupBox(
            "Customer Details"
        )

        customer_form = QFormLayout(
            customer_box
        )


        self.customer_name = QLineEdit()

        self.customer_gstin = QLineEdit()

        self.customer_address = QLineEdit()

        self.customer_mobile = QLineEdit()

        customer_form.addRow(
            "Customer Name",
            self.customer_name
        )

        customer_form.addRow(
            "GSTIN",
            self.customer_gstin
        )

        customer_form.addRow(
            "Address",
            self.customer_address
        )

        customer_form.addRow(
            "Mobile",
            self.customer_mobile
        )

        self.customer_name.textChanged.connect(
            lambda value: self.sync_buyer_field(
                self.invoice_customer_name,
                value
            )
        )

        self.customer_gstin.textChanged.connect(
            lambda value: self.sync_buyer_field(
                self.invoice_customer_gstin,
                value
            )
        )

        self.customer_address.textChanged.connect(
            lambda value: self.sync_buyer_field(
                self.invoice_customer_address,
                value
            )
        )

        self.customer_mobile.textChanged.connect(
            lambda value: self.sync_buyer_field(
                self.invoice_customer_mobile,
                value
            )
        )


        layout.addWidget(
            customer_box
        )


        # ----------------------------------------------------
        # GST
        # ----------------------------------------------------

        tax_box = QGroupBox(
            "Tax"
        )

        tax_form = QFormLayout(
            tax_box
        )


        self.gst_rate = NoWheelDoubleSpinBox()

        self.gst_rate.setRange(
            0,
            100
        )

        self.gst_rate.setDecimals(
            2
        )

        self.gst_rate.setValue(
            5
        )

        self.gst_rate.setSuffix(
            " %"
        )

        self.gst_rate.valueChanged.connect(
            self.refresh
        )


        tax_form.addRow(
            "GST Rate",
            self.spinbox_with_arrows(self.gst_rate)
        )


        layout.addWidget(
            tax_box
        )


        # ----------------------------------------------------
        # BUTTONS
        # ----------------------------------------------------

        buttons = QHBoxLayout()


        new_button = QPushButton(
            "New Invoice"
        )

        new_button.clicked.connect(
            self.new_invoice
        )


        save_button = QPushButton(
            "Save Invoice"
        )

        save_button.setObjectName(
            "primary"
        )

        save_button.clicked.connect(
            lambda checked=False: self.save_invoice()
        )


        pdf_button = QPushButton(
            "Download PDF"
        )

        pdf_button.setObjectName(
            "primary"
        )

        pdf_button.clicked.connect(
            self.export_pdf
        )


        buttons.addWidget(
            new_button
        )

        buttons.addWidget(
            save_button
        )

        buttons.addWidget(
            pdf_button
        )

        buttons.addStretch()


        layout.addLayout(
            buttons
        )

        layout.addStretch()


    # ========================================================
    # INVOICE TAB
    # ========================================================

    def create_legacy_invoice_tab(self):

        layout = QVBoxLayout(
            self.invoice_tab
        )

        buyer_box = QGroupBox(
            "Buyer Details"
        )

        buyer_form = QFormLayout(
            buyer_box
        )

        self.invoice_customer_name = QLineEdit()
        self.invoice_customer_gstin = QLineEdit()
        self.invoice_customer_address = QLineEdit()
        self.invoice_customer_mobile = QLineEdit()

        buyer_form.addRow(
            "Customer Name",
            self.invoice_customer_name
        )
        buyer_form.addRow(
            "GSTIN",
            self.invoice_customer_gstin
        )
        buyer_form.addRow(
            "Address",
            self.invoice_customer_address
        )
        buyer_form.addRow(
            "Mobile",
            self.invoice_customer_mobile
        )

        for main_field, invoice_field in (
            (self.customer_name, self.invoice_customer_name),
            (self.customer_gstin, self.invoice_customer_gstin),
            (self.customer_address, self.invoice_customer_address),
            (self.customer_mobile, self.invoice_customer_mobile),
        ):
            invoice_field.textChanged.connect(
                lambda value, target=main_field: self.sync_buyer_field(
                    target,
                    value
                )
            )

        layout.addWidget(
            buyer_box
        )


        box = QGroupBox(
            "Items"
        )

        box_layout = QVBoxLayout(
            box
        )


        self.items = QTableWidget(
            10,
            7
        )


        self.items.setHorizontalHeaderLabels([
            "No.",
            "Description of Goods",
            "HSN / SAC",
            "Quantity",
            "Rate",
            "Per",
            "Amount"
        ])


        self.items.horizontalHeader().setSectionResizeMode(
            1,
            QHeaderView.ResizeMode.Stretch
        )


        for column in (0, 2, 3, 4, 5, 6):

            self.items.horizontalHeader().setSectionResizeMode(
                column,
                QHeaderView.ResizeMode.ResizeToContents
            )

        for row in range(self.items.rowCount()):
            for column in range(self.items.columnCount()):
                item = QTableWidgetItem("")
                if column not in (1, 3, 4):
                    item.setFlags(
                        item.flags()
                        & ~Qt.ItemFlag.ItemIsEditable
                    )
                self.items.setItem(row, column, item)


        self.items.itemChanged.connect(
            self.refresh
        )


        box_layout.addWidget(
            self.items
        )


        layout.addWidget(
            box
        )


        # ----------------------------------------------------
        # TOTAL
        # ----------------------------------------------------

        total_box = QGroupBox(
            "Invoice Total"
        )

        total_layout = QVBoxLayout(
            total_box
        )


        self.summary = QLabel()

        self.summary.setFont(
            QFont(
                "Segoe UI",
                11,
                QFont.Weight.DemiBold
            )
        )


        total_layout.addWidget(
            self.summary
        )


        layout.addWidget(
            total_box
        )


        # ----------------------------------------------------
        # BUTTONS
        # ----------------------------------------------------

        buttons = QHBoxLayout()


        save_button = QPushButton(
            "Save Invoice"
        )

        save_button.setObjectName(
            "primary"
        )

        save_button.clicked.connect(
            lambda checked=False: self.save_invoice()
        )


        pdf_button = QPushButton(
            "Download PDF"
        )

        pdf_button.setObjectName(
            "primary"
        )

        pdf_button.clicked.connect(
            self.export_pdf
        )


        buttons.addWidget(
            save_button
        )

        buttons.addWidget(
            pdf_button
        )

        buttons.addStretch()


        layout.addLayout(
            buttons
        )


    def sync_buyer_field(self, target, value):

        if target.text() != value:
            target.setText(value)


    def create_invoice_tab(self):

        layout = QVBoxLayout(self.invoice_tab)
        self.create_mode_tabs = QTabWidget()
        self.create_mode_tabs.addTab(QWidget(), "+ Invoice")
        self.create_mode_tabs.addTab(QWidget(), "+ Quotation")
        self.create_mode_tabs.setFixedHeight(58)
        self.create_mode_tabs.setStyleSheet(
            "QTabWidget::pane { border: 0; }"
            "QTabBar::tab { min-width: 150px; padding: 9px 18px; "
            "color: #536873; background: #e9eff2; border: 0; }"
            "QTabBar::tab:selected { color: #155b46; background: #d0e6dd; "
            "font-weight: 600; border-bottom: 3px solid #21866b; }"
        )
        layout.addWidget(self.create_mode_tabs)
        scroll_area = QScrollArea()
        self.invoice_scroll_area = scroll_area
        scroll_area.setWidgetResizable(True)
        page = QWidget()
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(18, 18, 18, 18)

        self.invoice_request_history_box = QGroupBox("Payment Request History")
        history_layout = QVBoxLayout(self.invoice_request_history_box)
        history_actions = QHBoxLayout()
        self.open_invoice_request_button = QPushButton("Open PDF")
        self.open_invoice_request_button.clicked.connect(
            self.open_selected_invoice_request
        )
        self.download_invoice_request_button = QPushButton("Download PDF")
        self.download_invoice_request_button.clicked.connect(
            self.download_selected_invoice_request
        )
        self.send_paid_invoice_button = QPushButton("Send Paid PDF")
        self.send_paid_invoice_button.clicked.connect(
            self.send_selected_paid_invoice
        )
        self.open_invoice_request_button.setEnabled(False)
        self.download_invoice_request_button.setEnabled(False)
        self.send_paid_invoice_button.setEnabled(False)
        history_actions.addWidget(self.open_invoice_request_button)
        history_actions.addWidget(self.download_invoice_request_button)
        history_actions.addWidget(self.send_paid_invoice_button)
        history_actions.addStretch()
        history_layout.addLayout(history_actions)
        self.invoice_request_history_table = QTableWidget(0, 4)
        self.invoice_request_history_table.setHorizontalHeaderLabels(
            ["Date / Time", "Payment Request", "Status", "PDF"]
        )
        self.invoice_request_history_table.setEditTriggers(
            QTableWidget.EditTrigger.NoEditTriggers
        )
        self.invoice_request_history_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )
        self.invoice_request_history_table.setSelectionMode(
            QTableWidget.SelectionMode.SingleSelection
        )
        self.invoice_request_history_table.verticalHeader().setVisible(False)
        self.invoice_request_history_table.horizontalHeader().setSectionResizeMode(
            3, QHeaderView.ResizeMode.Stretch
        )
        self.invoice_request_history_table.setMinimumHeight(240)
        self.invoice_request_history_table.setMaximumHeight(360)
        self.invoice_request_history_table.itemSelectionChanged.connect(
            self.update_invoice_request_actions
        )
        self.invoice_request_history_table.cellDoubleClicked.connect(
            lambda row, column: self.open_selected_invoice_request()
        )
        history_layout.addWidget(self.invoice_request_history_table)
        self.invoice_request_history_box.setVisible(False)

        seller_box = QGroupBox("Seller Details")
        seller_grid = QGridLayout(seller_box)
        self.invoice_heading = QLineEdit("TAX INVOICE")
        self.seller_name = QLineEdit("ADOODS GARMENTS")
        self.seller_business = QLineEdit("ALL KINDS OF T-SHIRT MANUFACTURE")
        self.seller_address_1 = QLineEdit("88-1/50M, 1st FLOOR NEAR SRINIVASA THEATRE,")
        self.seller_address_2 = QLineEdit("ULAVARSANDHAI ROAD, HOSUR - 635 109")
        self.seller_gstin = QLineEdit("33DJRPP6135G1ZW")
        self.seller_state = QLineEdit("STATE: TAMIL NADU   CODE: 33")
        self.seller_website_phone = QLineEdit("Website: adoods.com   PHONE: +91 9514109080")
        self.seller_email_phone = QLineEdit("Email: adoodshosur@gmail.com   +91 6382208004")
        self.invoice_heading.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.invoice_heading.setStyleSheet("font: 20pt 'Arial Black';")
        self.seller_name.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.seller_name.setStyleSheet("font: bold 26pt 'Arial Black';")
        seller_rows = (
            ("Title", self.invoice_heading),
            ("Company", self.seller_name),
            ("Business", self.seller_business),
            ("Address", self.seller_address_1),
            ("Address", self.seller_address_2),
            ("GSTIN", self.seller_gstin),
            ("State", self.seller_state),
            ("Website / Phone", self.seller_website_phone),
            ("Email / Phone", self.seller_email_phone),
        )
        for row, (label, field) in enumerate(seller_rows):
            seller_grid.addWidget(QLabel(label), row, 0)
            seller_grid.addWidget(field, row, 1)
        page_layout.addWidget(seller_box)

        party_grid = QGridLayout()
        buyer_box = QGroupBox("Buyer Details")
        buyer_form = QFormLayout(buyer_box)
        self.customer_name = QLineEdit()
        self.customer_address = QLineEdit()
        self.customer_address_2 = QLineEdit()
        self.customer_mobile = QLineEdit()
        self.customer_mobile.setPlaceholderText("10-digit Indian mobile number")
        self.customer_mobile.setValidator(
            QRegularExpressionValidator(
                QRegularExpression(r"[+0-9() -]{0,20}"),
                self.customer_mobile,
            )
        )
        self.customer_gstin = QLineEdit()
        self.buyer_selector = NoWheelComboBox()
        self.buyer_selector.currentIndexChanged.connect(self.load_buyer_profile)
        buyer_profile_row = QHBoxLayout()
        buyer_profile_row.addWidget(self.buyer_selector, 1)
        new_buyer_button = QPushButton("New Buyer")
        new_buyer_button.clicked.connect(self.new_buyer_profile)
        save_buyer_button = QPushButton("Save / Update")
        save_buyer_button.clicked.connect(self.save_buyer_profile)
        delete_buyer_button = QPushButton("Delete")
        delete_buyer_button.clicked.connect(self.delete_buyer_profile)
        buyer_profile_row.addWidget(new_buyer_button)
        buyer_profile_row.addWidget(save_buyer_button)
        buyer_profile_row.addWidget(delete_buyer_button)
        buyer_form.addRow("Saved buyer details", buyer_profile_row)
        buyer_form.addRow("Name", self.customer_name)
        buyer_form.addRow("Address line 1", self.customer_address)
        buyer_form.addRow("Address line 2", self.customer_address_2)
        buyer_form.addRow("Mobile", self.customer_mobile)
        buyer_form.addRow("GSTIN", self.customer_gstin)

        invoice_box = QGroupBox("Invoice Details")
        self.document_details_box = invoice_box
        invoice_form = QFormLayout(invoice_box)
        self.invoice_no = QLineEdit()
        self.document_kind = NoWheelComboBox()
        self.document_kind.addItems(["Invoice", "Quotation"])
        self.document_kind.hide()
        self.create_mode_tabs.currentChanged.connect(
            self.document_kind.setCurrentIndex
        )
        self.document_number_label = QLabel("Invoice No")
        self.invoice_date = NoWheelDateEdit(QDate.currentDate())
        self.invoice_date.setCalendarPopup(True)
        self.invoice_date.setDisplayFormat("dd/MM/yyyy")
        self.invoice_date.setKeyboardTracking(False)
        self.po_number = QLineEdit()
        self.invoice_type = NoWheelComboBox()
        self.invoice_type.addItems(["IGST", "CGST + SGST"])
        self.gst_rate = NoWheelDoubleSpinBox()
        self.gst_rate.setRange(0, 100)
        self.gst_rate.setDecimals(2)
        self.gst_rate.setValue(5)
        self.gst_rate.setSuffix(" %")
        invoice_form.addRow(self.document_number_label, self.invoice_no)
        invoice_form.addRow("Date", self.invoice_date)
        invoice_form.addRow("P.O No", self.po_number)
        invoice_form.addRow("Tax Type", self.invoice_type)
        invoice_form.addRow("GST Rate", self.spinbox_with_arrows(self.gst_rate))
        self.document_kind.currentIndexChanged.connect(self.document_kind_changed)
        self.document_kind.currentIndexChanged.connect(
            self.create_mode_tabs.setCurrentIndex
        )
        self.invoice_no.editingFinished.connect(self.load_invoice)
        party_grid.addWidget(buyer_box, 0, 0)
        party_grid.addWidget(invoice_box, 0, 1)
        page_layout.addLayout(party_grid)

        self.payment_percent = NoWheelDoubleSpinBox()
        self.payment_percent.setRange(0, 100)
        self.payment_percent.setDecimals(2)
        self.payment_percent.setValue(50)
        self.payment_percent.setSuffix(" % of invoice")
        self.payment_percent.valueChanged.connect(self.refresh_payment_request)

        self.invoice_payments_box = QGroupBox("Invoice Receipts")
        invoice_payments_layout = QVBoxLayout(self.invoice_payments_box)
        receipt_form = QFormLayout()
        receipt_form.addRow(
            "Payment Request",
            self.spinbox_with_arrows(self.payment_percent, step_size=5),
        )
        self.payment_request_amount = QLabel("₹0.00")
        receipt_form.addRow("Request amount", self.payment_request_amount)
        self.remaining_percent_label = QLabel("Balance remaining: 100%")
        receipt_form.addRow("Invoice balance", self.remaining_percent_label)
        self.receipt_amount = AmountEntryDoubleSpinBox()
        self.receipt_amount.setRange(0, 999999999.99)
        self.receipt_amount.setDecimals(2)
        self.receipt_amount.setPrefix("₹ ")
        self._auto_receipt_amount = Decimal("0.00")
        self._receipt_amount_manual_override = False
        self._updating_receipt_amount = False
        self.receipt_amount.valueChanged.connect(
            self.track_receipt_amount_override
        )
        self.receipt_method = NoWheelComboBox()
        self.receipt_method.addItems(
            ["Cash", "UPI", "Bank Transfer", "Cheque", "Other"]
        )
        self.receipt_date = NoWheelDateEdit(QDate.currentDate())
        self.receipt_date.setCalendarPopup(True)
        self.receipt_date.setDisplayFormat("dd/MM/yyyy")
        self.receipt_reference = QLineEdit()
        self.receipt_reference.setPlaceholderText("Optional transaction / receipt reference")
        receipt_form.addRow(
            "Amount received", self.spinbox_with_arrows(self.receipt_amount)
        )
        receipt_form.addRow("Payment method", self.receipt_method)
        receipt_form.addRow("Payment date", self.receipt_date)
        receipt_form.addRow("Reference", self.receipt_reference)
        invoice_payments_layout.addLayout(receipt_form)
        self.record_payment_button = QPushButton("Record Payment")
        self.record_payment_button.clicked.connect(self.record_invoice_payment)
        self.record_payment_button.setMaximumWidth(160)
        self.record_payment_button.setMinimumHeight(34)
        self.edit_payment_button = QPushButton("Edit Selected")
        self.edit_payment_button.clicked.connect(self.edit_selected_invoice_payment)
        self.delete_payment_button = QPushButton("Delete Selected")
        self.delete_payment_button.clicked.connect(self.delete_selected_invoice_payment)
        self.edit_payment_button.setEnabled(False)
        self.delete_payment_button.setEnabled(False)
        payment_actions = QHBoxLayout()
        payment_actions.addWidget(self.record_payment_button)
        payment_actions.addWidget(self.edit_payment_button)
        payment_actions.addWidget(self.delete_payment_button)
        payment_actions.addStretch()
        self.invoice_payment_summary = QLabel("Save the invoice before recording payments.")
        payment_actions.addWidget(self.invoice_payment_summary, 1)
        invoice_payments_layout.addLayout(payment_actions)
        self.invoice_payments_table = QTableWidget(0, 4)
        self.invoice_payments_table.setHorizontalHeaderLabels(
            ["Date", "Amount", "Method", "Reference"]
        )
        self.invoice_payments_table.setEditTriggers(
            QTableWidget.EditTrigger.NoEditTriggers
        )
        self.invoice_payments_table.verticalHeader().setVisible(False)
        self.invoice_payments_table.horizontalHeader().setSectionResizeMode(
            3, QHeaderView.ResizeMode.Stretch
        )
        self.invoice_payments_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )
        self.invoice_payments_table.setSelectionMode(
            QTableWidget.SelectionMode.SingleSelection
        )
        self.invoice_payments_table.setMinimumHeight(160)
        self.invoice_payments_table.setMaximumHeight(240)
        self.invoice_payments_table.itemSelectionChanged.connect(
            self.update_invoice_payment_actions
        )
        invoice_payments_layout.addWidget(self.invoice_payments_table)
        self.record_payment_button.setEnabled(False)
        self.invoice_payments_box.setVisible(False)

        items_box = QGroupBox("Items")
        items_layout = QVBoxLayout(items_box)
        self.items = QTableWidget(0, 8)
        self.items.setFixedHeight(252)
        self.items.setWordWrap(True)
        self.items.setAlternatingRowColors(True)
        self.items.setStyleSheet(
            "QTableWidget { background: #ffffff; alternate-background-color: #f4f7fb; }"
            " QTableWidget::item { color: #1f2937; }"
        )
        self.items.verticalHeader().setDefaultSectionSize(52)
        self.items.verticalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Fixed
        )
        self.items.setHorizontalHeaderLabels([
            "SL.NO", "Description of Goods", "HSN/SAC", "Quantity",
            "Rate", "Per", "Amount", "Remove"
        ])
        self.items.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.Stretch
        )
        for column in (0, 2, 3, 4, 5, 6, 7):
            self.items.horizontalHeader().setSectionResizeMode(
                column, QHeaderView.ResizeMode.ResizeToContents
            )
        item_buttons = QHBoxLayout()
        add_item_button = QPushButton("Add Item")
        add_item_button.clicked.connect(self.add_item_row)
        item_buttons.addWidget(add_item_button)
        item_buttons.addStretch()
        items_layout.addLayout(item_buttons)
        items_layout.addWidget(self.items)
        page_layout.addWidget(items_box)

        totals_grid = QGridLayout()
        self.summary = QLabel()
        self.summary.setFont(QFont("Arial Black", 12, QFont.Weight.Bold))
        self.hsn_code = QLineEdit("6006")
        self.amount_words = QLineEdit()
        self.tax_words = QLineEdit()
        self.amount_words.setReadOnly(True)
        self.tax_words.setReadOnly(True)
        totals_grid.addWidget(QLabel("Amount Chargeable (in words)"), 0, 0)
        totals_grid.addWidget(self.amount_words, 0, 1, 1, 3)
        totals_grid.addWidget(QLabel("Tax Amount (in words)"), 1, 0)
        totals_grid.addWidget(self.tax_words, 1, 1, 1, 3)
        totals_grid.addWidget(QLabel("HSN/SAC"), 2, 0)
        totals_grid.addWidget(self.hsn_code, 2, 1)

        self.hsn_table = QTableWidget(2, 8)
        self.hsn_table.setHorizontalHeaderLabels([
            "HSN/SAC", "", "Taxable Value", "Tax Rate", "Tax Amount",
            "Tax Rate", "Tax Amount", "Total Tax"
        ])
        hsn_header = self.hsn_table.horizontalHeader()
        hsn_header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        hsn_header.setMinimumSectionSize(72)
        hsn_header.setMinimumHeight(58)
        hsn_header.setStretchLastSection(True)
        self.hsn_table.setWordWrap(True)
        self.hsn_table.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
        )
        self.hsn_table.setMinimumHeight(124)
        for column, width in enumerate((82, 48, 116, 82, 102, 82, 102, 102)):
            hsn_header.resizeSection(column, width)
        self.hsn_table.setEditTriggers(
            QTableWidget.EditTrigger.NoEditTriggers
        )
        self.hsn_table.verticalHeader().setVisible(False)
        totals_grid.addWidget(self.hsn_table, 3, 0, 1, 4)
        page_layout.addWidget(self.summary)
        page_layout.addLayout(totals_grid)

        bank_box = QGroupBox("Bank Profiles")
        self.bank_box = bank_box
        bank_grid = QGridLayout(bank_box)
        self.bank_name = QLineEdit("K V B")
        self.bank_account = QLineEdit("1607135000011387")
        self.bank_ifsc = QLineEdit("KVBL0001607")
        self.bank_upi_id = QLineEdit()
        self.bank_upi_id.setPlaceholderText("Required for PhonePe / Google Pay QR")
        self.bank_upi_label = QLabel("UPI ID for amount-specific QR")
        self.bank_selector = NoWheelComboBox()
        self.bank_selector.currentIndexChanged.connect(self.load_bank_profile)
        new_bank_button = QPushButton("New Bank")
        new_bank_button.clicked.connect(self.new_bank_profile)
        save_bank_button = QPushButton("Save / Update")
        save_bank_button.clicked.connect(self.save_bank_profile)
        delete_bank_button = QPushButton("Delete Selected")
        delete_bank_button.clicked.connect(self.delete_bank_profile)
        bank_profile_row = QHBoxLayout()
        bank_profile_row.addWidget(self.bank_selector, 1)
        bank_profile_row.addWidget(new_bank_button)
        bank_profile_row.addWidget(save_bank_button)
        bank_profile_row.addWidget(delete_bank_button)
        bank_grid.addWidget(QLabel("Saved profiles"), 0, 0)
        bank_grid.addLayout(bank_profile_row, 0, 1)
        for row, (label, field) in enumerate((
            ("Bank name", self.bank_name),
            ("Account number", self.bank_account),
            ("IFSC", self.bank_ifsc),
        ), start=1):
            bank_grid.addWidget(QLabel(label), row, 0)
            bank_grid.addWidget(field, row, 1)
        bank_grid.addWidget(self.bank_upi_label, 4, 0)
        bank_grid.addWidget(self.bank_upi_id, 4, 1)
        page_layout.addWidget(self.invoice_payments_box)
        page_layout.addWidget(bank_box)

        footer_box = QGroupBox("Invoice Footer")
        footer_grid = QGridLayout(footer_box)
        self.bundle_count = QLineEdit()
        self.through = QLineEdit()
        self.declaration = QTextEdit(
            "We declare that the Invoice shows the actual price of the goods "
            "described and that all particulars are true and correct."
        )
        self.declaration.setMinimumHeight(72)
        self.declaration.setMaximumHeight(96)
        self.seller_signatory = QLineEdit("for ADOODS GARMENTS")
        self.jurisdiction = QLineEdit(
            "SUBJECT TO HOSUR JURISDICTION. THIS IS COMPUTER GENERATED INVOICE"
        )
        for field in self.seller_profile_fields():
            field.editingFinished.connect(self.save_seller_profile)
        self.signature_image_path = (
            SIGNATURE_PATH if SIGNATURE_PATH.is_file() else None
        )
        footer_fields = (
            ("No of Bundles", self.bundle_count),
            ("Through", self.through),
            ("Declaration", self.declaration),
            ("Authorised Signatory", self.seller_signatory),
            ("Footer", self.jurisdiction),
        )
        for row, (label, field) in enumerate(footer_fields):
            footer_grid.addWidget(QLabel(label), row, 0)
            footer_grid.addWidget(field, row, 1)
        signature_row = QHBoxLayout()
        self.signature_preview = QLabel("No signature")
        self.signature_preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.signature_preview.setFixedSize(148, 54)
        self.signature_preview.setStyleSheet(
            "QLabel { background: #f5f8f7; color: #71817b; "
            "border: 1px dashed #9db4aa; border-radius: 4px; }"
        )
        self.signature_image_label = QLabel(
            "Signature ready for invoices"
            if self.signature_image_path else "No signature uploaded"
        )
        self.signature_image_label.setStyleSheet("color: #657780;")
        self.upload_signature_button = QPushButton(
            "Replace signature" if self.signature_image_path else "Upload signature PNG"
        )
        self.upload_signature_button.clicked.connect(self.upload_signature)
        signature_row.addWidget(self.signature_preview)
        signature_row.addWidget(self.signature_image_label)
        signature_row.addWidget(self.upload_signature_button)
        signature_row.addStretch()
        footer_grid.addLayout(signature_row, len(footer_fields), 0, 1, 2)
        self.refresh_signature_display()
        page_layout.addWidget(footer_box)

        action_row = QHBoxLayout()
        self.new_document_button = QPushButton("New Invoice")
        self.new_document_button.clicked.connect(self.new_invoice)
        self.whatsapp_button = QPushButton("WhatsApp Invoice")
        self.whatsapp_button.setObjectName("primary")
        self.whatsapp_button.clicked.connect(self.share_document_whatsapp)
        self.whatsapp_button.setVisible(False)
        self.attach_invoice_checkbox = QCheckBox("Attach invoice when sharing")
        self.attach_invoice_checkbox.setChecked(True)
        self.attach_invoice_checkbox.setVisible(False)
        self.save_document_button = QPushButton("Save Invoice")
        self.save_document_button.setObjectName("primary")
        self.save_document_button.clicked.connect(
            lambda checked=False: self.save_document_bundle()
        )
        self.private_invoice_checkbox = QCheckBox("Private invoice")
        action_row.addWidget(self.new_document_button)
        action_row.addWidget(self.private_invoice_checkbox)
        action_row.addWidget(self.save_document_button)
        action_row.addWidget(self.whatsapp_button)
        action_row.addWidget(self.attach_invoice_checkbox)
        action_row.addStretch()
        page_layout.addLayout(action_row)
        page_layout.addWidget(self.invoice_request_history_box)

        scroll_area.setWidget(page)
        layout.addWidget(scroll_area)
        self.items.itemChanged.connect(self.refresh)
        self.invoice_type.currentIndexChanged.connect(self.refresh)
        self.invoice_date.dateChanged.connect(self.refresh)
        self.gst_rate.valueChanged.connect(self.refresh)
        self.load_seller_profile()
        self.refresh_buyer_profiles()
        self.refresh_bank_profiles()
        self.refresh_invoice_history()
        self.document_kind_changed(self.document_kind.currentIndex())


    def create_pending_invoices_tab(self):

        self.pending_invoices_tab = QWidget()
        layout = QVBoxLayout(self.pending_invoices_tab)
        toolbar = QHBoxLayout()
        self.pending_invoice_search = QLineEdit()
        self.pending_invoice_search.setPlaceholderText(
            "Search customer or invoice number"
        )
        self.pending_invoice_search.textChanged.connect(
            lambda value: self.reset_pending_invoice_page()
        )
        toolbar.addWidget(self.pending_invoice_search, 1)
        self.include_paid_invoices = QCheckBox("Include paid invoices")
        self.include_paid_invoices.toggled.connect(
            lambda checked: self.reset_pending_invoice_page()
        )
        toolbar.addWidget(self.include_paid_invoices)
        self.open_pending_invoice_button = QPushButton("Open Selected Invoice")
        self.open_pending_invoice_button.clicked.connect(
            self.open_selected_pending_invoice
        )
        toolbar.addWidget(self.open_pending_invoice_button)
        layout.addLayout(toolbar)

        self.pending_invoices_table = QTableWidget(0, 6)
        self.pending_invoices_table.setHorizontalHeaderLabels(
            ["Invoice No", "Date", "Customer", "Total", "Received", "Balance Due"]
        )
        self.pending_invoices_table.setEditTriggers(
            QTableWidget.EditTrigger.NoEditTriggers
        )
        self.pending_invoices_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )
        self.pending_invoices_table.setSelectionMode(
            QTableWidget.SelectionMode.SingleSelection
        )
        self.pending_invoices_table.verticalHeader().setVisible(False)
        self.pending_invoices_table.horizontalHeader().setSectionResizeMode(
            2, QHeaderView.ResizeMode.Stretch
        )
        self.pending_invoices_table.cellDoubleClicked.connect(
            lambda row, column: self.open_selected_pending_invoice()
        )
        layout.addWidget(self.pending_invoices_table, 1)

        pagination = QHBoxLayout()
        self.pending_previous_button = QPushButton("Previous")
        self.pending_previous_button.clicked.connect(
            lambda checked=False: self.change_pending_invoice_page(-1)
        )
        self.pending_page_label = QLabel("Page 1 of 1")
        self.pending_page_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.pending_next_button = QPushButton("Next")
        self.pending_next_button.clicked.connect(
            lambda checked=False: self.change_pending_invoice_page(1)
        )
        pagination.addWidget(self.pending_previous_button)
        pagination.addWidget(self.pending_page_label, 1)
        pagination.addWidget(self.pending_next_button)
        layout.addLayout(pagination)

        self.pending_invoice_page = 0
        self.pending_invoice_page_size = 20
        self.refresh_pending_invoices()


    def reset_pending_invoice_page(self):

        self.pending_invoice_page = 0
        self.refresh_pending_invoices()


    def change_pending_invoice_page(self, offset):

        self.pending_invoice_page += offset
        self.refresh_pending_invoices()


    def refresh_pending_invoices(self):

        if not hasattr(self, "pending_invoices_table"):
            return

        search = self.pending_invoice_search.text().strip()
        search_pattern = f"%{search}%"
        paid_filter = "" if self.include_paid_invoices.isChecked() else (
            "COALESCE(i.grand_total, 0) - COALESCE(SUM(p.amount), 0) > 0.005 AND "
        )
        with sqlite3.connect(DB_PATH) as con:
            rows = con.execute(
                f"""
                SELECT i.invoice_no, i.invoice_date, i.customer_name,
                       i.grand_total, COALESCE(SUM(p.amount), 0) AS received
                FROM invoices AS i
                LEFT JOIN invoice_payments AS p
                  ON p.document_kind = i.document_kind
                 AND p.invoice_no = i.invoice_no
                WHERE i.document_kind = 'INVOICE'
                GROUP BY i.id
                HAVING {paid_filter}
                   (COALESCE(i.customer_name, '') LIKE ?
                    OR CAST(i.invoice_no AS TEXT) LIKE ?)
                ORDER BY i.invoice_date DESC, i.invoice_no DESC
                """,
                (search_pattern, search_pattern),
            ).fetchall()

        page_count = max(
            1,
            (len(rows) + self.pending_invoice_page_size - 1)
            // self.pending_invoice_page_size,
        )
        self.pending_invoice_page = min(
            max(0, self.pending_invoice_page), page_count - 1
        )
        start = self.pending_invoice_page * self.pending_invoice_page_size
        page_rows = rows[start:start + self.pending_invoice_page_size]
        self.pending_invoices_table.setRowCount(len(page_rows))
        for row_index, (invoice_no, invoice_date, customer, total, received) in enumerate(page_rows):
            balance_due = max(Decimal("0.00"), money(Decimal(str(total or 0)) - Decimal(str(received or 0))))
            values = (
                str(invoice_no),
                QDate.fromString(invoice_date, "yyyy-MM-dd").toString("dd/MM/yyyy"),
                customer or "Unnamed buyer",
                f"₹{money(total or 0):,.2f}",
                f"₹{money(received or 0):,.2f}",
                f"₹{balance_due:,.2f}",
            )
            for column, value in enumerate(values):
                self.pending_invoices_table.setItem(
                    row_index, column, QTableWidgetItem(value)
                )

        self.pending_page_label.setText(
            f"Page {self.pending_invoice_page + 1} of {page_count} · {len(rows)} invoices"
        )
        self.pending_previous_button.setEnabled(self.pending_invoice_page > 0)
        self.pending_next_button.setEnabled(self.pending_invoice_page + 1 < page_count)
        self.open_pending_invoice_button.setEnabled(bool(page_rows))


    def open_selected_pending_invoice(self):

        row = self.pending_invoices_table.currentRow()
        if row < 0:
            return

        invoice_number = self.pending_invoices_table.item(row, 0)
        if not invoice_number:
            return

        self.document_kind.setCurrentIndex(0)
        self.invoice_no.setText(invoice_number.text())
        self.load_invoice()
        self.main_tabs.setCurrentWidget(self.invoice_tab)


    def create_quotation_register_tab(self):

        self.quotation_register_tab = QWidget()
        layout = QVBoxLayout(self.quotation_register_tab)
        toolbar = QHBoxLayout()
        self.quotation_search = QLineEdit()
        self.quotation_search.setPlaceholderText(
            "Search customer or quotation number"
        )
        self.quotation_search.textChanged.connect(
            lambda value: self.reset_quotation_page()
        )
        toolbar.addWidget(self.quotation_search, 1)
        self.open_quotation_button = QPushButton("Open Selected Quotation")
        self.open_quotation_button.clicked.connect(self.open_selected_quotation)
        toolbar.addWidget(self.open_quotation_button)
        self.convert_selected_quotation_button = QPushButton(
            "Order Confirmed · Add to Invoice"
        )
        self.convert_selected_quotation_button.clicked.connect(
            self.convert_selected_quotation_to_invoice
        )
        self.convert_selected_quotation_button.setEnabled(False)
        toolbar.addWidget(self.convert_selected_quotation_button)
        layout.addLayout(toolbar)

        self.quotation_table = QTableWidget(0, 4)
        self.quotation_table.setHorizontalHeaderLabels(
            ["Quotation No", "Date", "Customer", "Total"]
        )
        self.quotation_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.quotation_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )
        self.quotation_table.setSelectionMode(
            QTableWidget.SelectionMode.SingleSelection
        )
        self.quotation_table.itemSelectionChanged.connect(
            self.update_quotation_actions
        )
        self.quotation_table.verticalHeader().setVisible(False)
        self.quotation_table.horizontalHeader().setSectionResizeMode(
            2, QHeaderView.ResizeMode.Stretch
        )
        self.quotation_table.cellDoubleClicked.connect(
            lambda row, column: self.open_selected_quotation()
        )
        layout.addWidget(self.quotation_table, 1)

        pagination = QHBoxLayout()
        self.quotation_previous_button = QPushButton("Previous")
        self.quotation_previous_button.clicked.connect(
            lambda checked=False: self.change_quotation_page(-1)
        )
        self.quotation_page_label = QLabel("Page 1 of 1")
        self.quotation_page_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.quotation_next_button = QPushButton("Next")
        self.quotation_next_button.clicked.connect(
            lambda checked=False: self.change_quotation_page(1)
        )
        pagination.addWidget(self.quotation_previous_button)
        pagination.addWidget(self.quotation_page_label, 1)
        pagination.addWidget(self.quotation_next_button)
        layout.addLayout(pagination)

        self.quotation_page = 0
        self.quotation_page_size = 20
        self.refresh_quotation_register()


    def reset_quotation_page(self):

        self.quotation_page = 0
        self.refresh_quotation_register()


    def change_quotation_page(self, offset):

        self.quotation_page += offset
        self.refresh_quotation_register()


    def refresh_quotation_register(self):

        if not hasattr(self, "quotation_table"):
            return

        search_pattern = f"%{self.quotation_search.text().strip()}%"
        with sqlite3.connect(DB_PATH) as con:
            rows = con.execute(
                """
                SELECT invoice_no, invoice_date, customer_name, grand_total
                FROM invoices
                WHERE document_kind = 'QUOTATION'
                  AND (COALESCE(customer_name, '') LIKE ?
                       OR CAST(invoice_no AS TEXT) LIKE ?)
                ORDER BY invoice_date DESC, invoice_no DESC
                """,
                (search_pattern, search_pattern),
            ).fetchall()

        page_count = max(
            1,
            (len(rows) + self.quotation_page_size - 1)
            // self.quotation_page_size,
        )
        self.quotation_page = min(max(0, self.quotation_page), page_count - 1)
        start = self.quotation_page * self.quotation_page_size
        page_rows = rows[start:start + self.quotation_page_size]
        self.quotation_table.setRowCount(len(page_rows))
        for row_index, (number, date, customer, total) in enumerate(page_rows):
            values = (
                str(number),
                QDate.fromString(date, "yyyy-MM-dd").toString("dd/MM/yyyy"),
                customer or "Unnamed buyer",
                f"₹{money(total or 0):,.2f}",
            )
            for column, value in enumerate(values):
                self.quotation_table.setItem(
                    row_index, column, QTableWidgetItem(value)
                )

        self.quotation_page_label.setText(
            f"Page {self.quotation_page + 1} of {page_count} · {len(rows)} quotations"
        )
        self.quotation_previous_button.setEnabled(self.quotation_page > 0)
        self.quotation_next_button.setEnabled(self.quotation_page + 1 < page_count)
        self.update_quotation_actions()


    def update_quotation_actions(self):

        selected = self.quotation_table.currentRow() >= 0
        self.open_quotation_button.setEnabled(selected)
        self.convert_selected_quotation_button.setEnabled(selected)


    def open_selected_quotation(self):

        row = self.quotation_table.currentRow()
        if row < 0:
            return
        quotation_number = self.quotation_table.item(row, 0)
        if not quotation_number:
            return

        self.document_kind.setCurrentIndex(1)
        self.invoice_no.setText(quotation_number.text())
        self.load_invoice()
        self.main_tabs.setCurrentWidget(self.invoice_tab)


    def convert_selected_quotation_to_invoice(self):

        row = self.quotation_table.currentRow()
        if row < 0:
            return
        quotation_number = self.quotation_table.item(row, 0)
        if not quotation_number:
            return

        self.document_kind.setCurrentIndex(1)
        self.invoice_no.setText(quotation_number.text())
        self.load_invoice()
        if self.loaded_invoice_no != int(quotation_number.text()):
            return
        self.main_tabs.setCurrentWidget(self.invoice_tab)
        self.convert_quotation_to_invoice()


    def convert_quotation_to_invoice(self):

        if document_kind_for_selection(self.document_kind.currentIndex()) != "QUOTATION":
            return

        quotation_state = self.capture_document_draft()
        if not quotation_state["items"]:
            QMessageBox.warning(
                self, "Quotation Has No Items", "Add at least one item before creating an invoice."
            )
            return

        self.save_document_draft("QUOTATION")
        self.document_kind.setCurrentIndex(0)
        quotation_state.update({
            "invoice_no": str(self.next_invoice_number("INVOICE")),
            "invoice_date": QDate.currentDate().toString("yyyy-MM-dd"),
            "payment_percent": 50,
            "private_invoice": False,
        })
        self.reset_receipt_amount_autofill()
        self.loaded_invoice_no = None
        self.apply_document_draft_state(quotation_state)
        self.save_document_draft("INVOICE")
        self.refresh_invoice_payments()
        self.refresh_invoice_request_history()
        self.main_tabs.setCurrentWidget(self.invoice_tab)
        self.invoice_scroll_area.verticalScrollBar().setValue(0)
        self.statusBar().showMessage(
            f"Quotation confirmed. Invoice {self.invoice_no.text()} is ready to edit."
        )


    def capture_document_draft(self):

        return {
            "invoice_no": self.invoice_no.text(),
            "invoice_date": self.invoice_date.date().toString("yyyy-MM-dd"),
            "customer_name": self.customer_name.text(),
            "customer_address": self.customer_address.text(),
            "customer_address_2": self.customer_address_2.text(),
            "customer_mobile": self.customer_mobile.text(),
            "customer_gstin": self.customer_gstin.text(),
            "buyer_profile": self.buyer_selector.currentData() or "",
            "po_number": self.po_number.text(),
            "invoice_type": self.invoice_type.currentIndex(),
            "gst_rate": self.gst_rate.value(),
            "items": [
                {
                    "description": self.cell(row, 1),
                    "hsn": self.cell(row, 2),
                    "quantity": self.cell(row, 3),
                    "rate": self.cell(row, 4),
                    "per": self.cell(row, 5),
                }
                for row in range(self.items.rowCount())
                if self.cell(row, 1)
            ],
            "hsn_code": self.hsn_code.text(),
            "bundle_count": self.bundle_count.text(),
            "through": self.through.text(),
            "bank_name": self.bank_name.text(),
            "bank_account": self.bank_account.text(),
            "bank_ifsc": self.bank_ifsc.text(),
            "bank_upi_id": self.bank_upi_id.text(),
            "bank_profile": self.bank_selector.currentData() or "",
            "payment_percent": self.payment_percent.value(),
            "declaration": self.declaration.toPlainText(),
            "seller_signatory": self.seller_signatory.text(),
            "jurisdiction": self.jurisdiction.text(),
            "private_invoice": self.private_invoice_checkbox.isChecked(),
            "custom_output_dir": str(self.custom_output_dir or ""),
        }


    def save_document_draft(self, document_kind):

        state_json = json.dumps(self.capture_document_draft())
        with sqlite3.connect(DB_PATH) as con:
            con.execute(
                """
                INSERT INTO document_drafts (document_kind, state_json)
                VALUES (?, ?)
                ON CONFLICT(document_kind) DO UPDATE SET
                    state_json = excluded.state_json
                """,
                (document_kind, state_json),
            )


    def delete_document_draft(self, document_kind):

        with sqlite3.connect(DB_PATH) as con:
            con.execute(
                "DELETE FROM document_drafts WHERE document_kind = ?",
                (document_kind,),
            )


    def restore_document_draft(self, document_kind):

        with sqlite3.connect(DB_PATH) as con:
            row = con.execute(
                "SELECT state_json FROM document_drafts WHERE document_kind = ?",
                (document_kind,),
            ).fetchone()
        if not row:
            return False

        state = json.loads(row[0])
        return self.apply_document_draft_state(state)


    def apply_document_draft_state(self, state):

        self.invoice_no.setText(state.get("invoice_no", ""))
        self.invoice_date.setDate(
            QDate.fromString(state.get("invoice_date", ""), "yyyy-MM-dd")
        )
        for key, field in (
            ("customer_name", self.customer_name),
            ("customer_address", self.customer_address),
            ("customer_address_2", self.customer_address_2),
            ("customer_mobile", self.customer_mobile),
            ("customer_gstin", self.customer_gstin),
            ("po_number", self.po_number),
            ("hsn_code", self.hsn_code),
            ("bundle_count", self.bundle_count),
            ("through", self.through),
            ("bank_name", self.bank_name),
            ("bank_account", self.bank_account),
            ("bank_ifsc", self.bank_ifsc),
            ("bank_upi_id", self.bank_upi_id),
            ("seller_signatory", self.seller_signatory),
            ("jurisdiction", self.jurisdiction),
        ):
            field.setText(state.get(key, ""))
        self.buyer_selector.blockSignals(True)
        self.buyer_selector.setCurrentIndex(
            max(0, self.buyer_selector.findData(state.get("buyer_profile", "")))
        )
        self.buyer_selector.blockSignals(False)
        self.bank_selector.blockSignals(True)
        self.bank_selector.setCurrentIndex(
            max(0, self.bank_selector.findData(state.get("bank_profile", "")))
        )
        self.bank_selector.blockSignals(False)
        self.invoice_type.setCurrentIndex(state.get("invoice_type", 0))
        self.gst_rate.setValue(state.get("gst_rate", 5))
        self.payment_percent.setValue(state.get("payment_percent", 50))
        self.declaration.setPlainText(state.get("declaration", ""))
        self.private_invoice_checkbox.setChecked(
            state.get("private_invoice", False)
        )
        custom_output_dir = state.get("custom_output_dir", "")
        self.custom_output_dir = Path(custom_output_dir) if custom_output_dir else None

        self.items.setRowCount(0)
        for item_data in state.get("items", []):
            self.add_item_row()
            row_index = self.items.rowCount() - 1
            for column, key in (
                (1, "description"), (2, "hsn"), (3, "quantity"),
                (4, "rate"), (5, "per"),
            ):
                self.items.item(row_index, column).setText(item_data.get(key, ""))
        self.refresh()
        self.refresh_invoice_request_history()
        return True


    def refresh_invoice_history(self):

        if not hasattr(self, "invoice_request_history_table"):
            return

        self.prune_missing_invoice_records()
        self.refresh_invoice_request_history()


    def refresh_invoice_request_history(self):

        if not hasattr(self, "invoice_request_history_table"):
            return

        self.invoice_request_history_table.setRowCount(0)
        is_invoice = (
            document_kind_for_selection(self.document_kind.currentIndex()) == "INVOICE"
        )
        self.invoice_request_history_box.setVisible(is_invoice)
        if not is_invoice:
            self.update_invoice_request_actions()
            return

        try:
            invoice_no = int(self.invoice_no.text())
        except ValueError:
            self.update_invoice_request_actions()
            return

        with sqlite3.connect(DB_PATH) as con:
            requests = con.execute(
                """
                  SELECT id, requested_at, request_percent, request_amount,
                      pdf_path, payment_update_path, final_pdf_path
                FROM invoice_requests
                WHERE invoice_no = ?
                ORDER BY requested_at DESC, id DESC
                """
            , (invoice_no,)).fetchall()
            invoice = con.execute(
                "SELECT grand_total FROM invoices "
                "WHERE invoice_no = ? AND document_kind = 'INVOICE'",
                (invoice_no,),
            ).fetchone()
            paid_total = con.execute(
                "SELECT COALESCE(SUM(amount), 0) FROM invoice_payments "
                "WHERE invoice_no = ? AND document_kind = 'INVOICE'",
                (invoice_no,),
            ).fetchone()[0]

        fully_paid = bool(
            invoice
            and money(paid_total) >= money(invoice[0] or 0)
            and money(invoice[0] or 0) > 0
        )
        received_percent = (
            money(paid_total) / money(invoice[0]) * 100
            if invoice and money(invoice[0] or 0) > 0 else Decimal("0")
        )

        for row, (
            request_id, requested_at, percent, amount, original_pdf_path,
            payment_update_path, final_pdf_path,
        ) in enumerate(requests):
            requested_time = datetime.fromisoformat(requested_at)
            request_status = (
                "Paid 100%" if fully_paid and row == 0
                else f"Received {received_percent:g}%" if row == 0 and paid_total > 0
                else "Earlier request" if row > 0
                else "Outstanding"
            )
            display_pdf_path = final_pdf_path or payment_update_path or original_pdf_path
            values = (
                requested_time.strftime("%d/%m/%Y %H:%M:%S"),
                f"{percent:g}% · ₹{money(amount):,.2f}",
                request_status,
                Path(display_pdf_path).name,
            )
            self.invoice_request_history_table.insertRow(row)
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                if column == 0:
                    item.setData(Qt.ItemDataRole.UserRole, request_id)
                if fully_paid and row == 0 and column == 2:
                    item.setBackground(QBrush(QColor("#dff2e3")))
                    item.setForeground(QBrush(QColor("#245c37")))
                self.invoice_request_history_table.setItem(row, column, item)
        self.update_invoice_request_actions()


    def update_invoice_request_actions(self):

        row = self.invoice_request_history_table.currentRow()
        selected = row >= 0
        self.open_invoice_request_button.setEnabled(selected)
        self.download_invoice_request_button.setEnabled(selected)
        is_paid_request = (
            selected
            and self.invoice_request_history_table.item(row, 2).text() == "Paid 100%"
        )
        self.send_paid_invoice_button.setEnabled(is_paid_request)


    def selected_invoice_request_path(self):

        row = self.invoice_request_history_table.currentRow()
        if row < 0:
            return None
        request_id = self.invoice_request_history_table.item(
            row, 0
        ).data(Qt.ItemDataRole.UserRole)
        with sqlite3.connect(DB_PATH) as con:
            result = con.execute(
                """
                SELECT COALESCE(
                    NULLIF(final_pdf_path, ''),
                    NULLIF(payment_update_path, ''),
                    pdf_path
                )
                FROM invoice_requests
                """
                "WHERE id = ? AND invoice_no = ?",
                (request_id, int(self.invoice_no.text())),
            ).fetchone()
        path = Path(result[0]) if result else None
        return path if path and path.is_file() else None


    def open_selected_invoice_request(self):

        pdf_path = self.selected_invoice_request_path()
        if not pdf_path:
            QMessageBox.warning(self, "PDF Missing", "This request PDF could not be found.")
            return
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(pdf_path.resolve()))):
            QMessageBox.warning(
                self, "PDF Could Not Be Opened", f"Open this PDF manually:\n{pdf_path}"
            )


    def download_selected_invoice_request(self):

        pdf_path = self.selected_invoice_request_path()
        if not pdf_path:
            QMessageBox.warning(self, "PDF Missing", "This request PDF could not be found.")
            return
        destination, _ = QFileDialog.getSaveFileName(
            self, "Download Request PDF", str(Path.home() / pdf_path.name),
            "PDF files (*.pdf)",
        )
        if not destination:
            return
        try:
            shutil.copy2(pdf_path, destination)
        except OSError as error:
            QMessageBox.warning(self, "Download Failed", str(error))
            return
        QMessageBox.information(self, "PDF Downloaded", destination)


    def send_selected_paid_invoice(self):

        if self.document_kind.currentIndex() != 0:
            return
        pdf_path = self.selected_invoice_request_path()
        if not pdf_path:
            QMessageBox.warning(self, "PDF Missing", "This paid invoice PDF could not be found.")
            return

        invoice_no = int(self.invoice_no.text())
        with sqlite3.connect(DB_PATH) as con:
            invoice = con.execute(
                "SELECT grand_total, customer_name FROM invoices "
                "WHERE invoice_no = ? AND document_kind = 'INVOICE'",
                (invoice_no,),
            ).fetchone()
            paid_total = con.execute(
                "SELECT COALESCE(SUM(amount), 0) FROM invoice_payments "
                "WHERE invoice_no = ? AND document_kind = 'INVOICE'",
                (invoice_no,),
            ).fetchone()[0]
        if not invoice or money(paid_total) < money(invoice[0] or 0):
            QMessageBox.warning(
                self, "Invoice Not Fully Paid", "Only a completed invoice can be sent from this action."
            )
            return

        mobile = normalize_indian_mobile(self.customer_mobile.text())
        if not mobile:
            QMessageBox.warning(
                self,
                "Invalid Buyer Mobile",
                "Enter a valid Indian mobile number before sharing.",
            )
            return

        message = "\n".join((
            f"Hello {invoice[1] or 'Customer'},",
            "",
            f"Thank you for choosing {self.seller_name.text()} for your garments order.",
            f"Please find attached the final paid invoice {invoice_no} for your records.",
            "",
            f"Invoice total: INR {money(invoice[0]):.2f}",
            f"Received: INR {money(paid_total):.2f}",
            "Balance due: INR 0.00",
            "",
            "We appreciate your business.",
            "Regards,",
            self.seller_name.text(),
        ))
        mime_data = QMimeData()
        mime_data.setUrls([QUrl.fromLocalFile(str(pdf_path.resolve()))])
        QApplication.clipboard().setMimeData(mime_data)
        opened = self.open_whatsapp_chat(mobile, message)
        if opened:
            QMessageBox.information(
                self,
                "Attach Paid Invoice",
                f"The completed invoice PDF is copied. Paste it into the WhatsApp chat (Ctrl+V):\n\n{pdf_path.name}",
            )
        else:
            QMessageBox.warning(
                self,
                "WhatsApp Could Not Be Opened",
                f"Open WhatsApp and paste the completed invoice PDF (Ctrl+V):\n{pdf_path.name}",
            )


    def refresh_invoice_payments(self):

        if not hasattr(self, "invoice_payments_table"):
            return

        self.invoice_payments_table.setRowCount(0)
        if document_kind_for_selection(self.document_kind.currentIndex()) != "INVOICE":
            self.record_payment_button.setEnabled(False)
            self.invoice_payment_summary.setText(
                "Payments are recorded against invoices, not quotations."
            )
            self.refresh_payment_request()
            return

        try:
            invoice_no = int(self.invoice_no.text())
        except ValueError:
            self.record_payment_button.setEnabled(False)
            self.invoice_payment_summary.setText(
                "Save the invoice before recording payments."
            )
            self.refresh_payment_request()
            return

        with sqlite3.connect(DB_PATH) as con:
            invoice = con.execute(
                "SELECT grand_total FROM invoices "
                "WHERE invoice_no = ? AND document_kind = 'INVOICE'",
                (invoice_no,),
            ).fetchone()
            payments = con.execute(
                """
                SELECT id, paid_on, amount, method, reference
                FROM invoice_payments
                WHERE document_kind = 'INVOICE' AND invoice_no = ?
                ORDER BY paid_on, id
                """,
                (invoice_no,),
            ).fetchall()

        if not invoice:
            self.record_payment_button.setEnabled(False)
            self.invoice_payment_summary.setText(
                "Save the invoice before recording payments."
            )
            self.refresh_payment_request()
            return

        grand_total = money(invoice[0] or 0)
        paid_total = money(
            sum((Decimal(str(payment[2])) for payment in payments), Decimal("0"))
        )
        balance_due = max(Decimal("0.00"), money(grand_total - paid_total))
        for row, (payment_id, paid_on, amount, method, reference) in enumerate(payments):
            self.invoice_payments_table.insertRow(row)
            date_text = QDate.fromString(paid_on, "yyyy-MM-dd").toString("dd/MM/yyyy")
            for column, value in enumerate((
                date_text or paid_on,
                f"₹{money(amount):,.2f}",
                method,
                reference,
            )):
                self.invoice_payments_table.setItem(
                    row, column, QTableWidgetItem(str(value))
                )
            self.invoice_payments_table.item(row, 0).setData(
                Qt.ItemDataRole.UserRole, payment_id
            )

        self.invoice_payment_summary.setText(
            f"Received: ₹{paid_total:,.2f}    Balance due: ₹{balance_due:,.2f}"
        )
        self.record_payment_button.setEnabled(balance_due > 0)
        self.update_invoice_payment_actions()
        self.refresh_payment_request()


    def refresh_payment_request(self):

        if not hasattr(self, "payment_request_amount"):
            return

        total = self.calculate()[3] if hasattr(self, "items") else Decimal("0.00")
        paid_total = Decimal("0.00")
        if (
            document_kind_for_selection(self.document_kind.currentIndex()) == "INVOICE"
            and self.invoice_no.text().strip().isdigit()
        ):
            paid_total = self.invoice_paid_total(int(self.invoice_no.text()))

        balance_due = max(Decimal("0.00"), money(total - paid_total))
        if total > 0 and balance_due == 0 and self.payment_percent.value() != 0:
            self.payment_percent.setValue(0)
        requested_amount = min(
            money(total * Decimal(str(self.payment_percent.value())) / 100),
            balance_due,
        )
        remaining_percent = (
            balance_due / total * 100 if total > 0 else Decimal("0")
        )
        self.payment_request_amount.setText(f"₹{requested_amount:,.2f}")
        self.remaining_percent_label.setText(
            f"Balance remaining: {remaining_percent:.2f}%"
        )
        current_amount = money(self.receipt_amount.value())
        if (
            not self._receipt_amount_manual_override
            or current_amount == self._auto_receipt_amount
        ):
            self._updating_receipt_amount = True
            self.receipt_amount.setValue(float(requested_amount))
            self._updating_receipt_amount = False
            self._auto_receipt_amount = requested_amount


    def track_receipt_amount_override(self, value):

        if self._updating_receipt_amount:
            return
        if money(value) != self._auto_receipt_amount:
            self._receipt_amount_manual_override = True


    def reset_receipt_amount_autofill(self):

        self._receipt_amount_manual_override = False
        self._auto_receipt_amount = Decimal("0.00")
        self._updating_receipt_amount = True
        self.receipt_amount.setValue(0)
        self._updating_receipt_amount = False


    def invoice_paid_total(self, invoice_no):

        with sqlite3.connect(DB_PATH) as con:
            paid_total = con.execute(
                """
                SELECT COALESCE(SUM(amount), 0)
                FROM invoice_payments
                WHERE document_kind = 'INVOICE' AND invoice_no = ?
                """,
                (invoice_no,),
            ).fetchone()[0]
        return money(paid_total)


    def update_invoice_payment_actions(self):

        has_selection = self.invoice_payments_table.currentRow() >= 0
        self.edit_payment_button.setEnabled(has_selection)
        self.delete_payment_button.setEnabled(has_selection)


    def refresh_invoice_after_payment_change(self):

        if not self.save_invoice(show_dialog=False, show_success=False):
            QMessageBox.warning(
                self,
                "Invoice Refresh Failed",
                "The receipt was changed, but the invoice workbook could not be refreshed.",
            )
            return False
        if not self.export_pdf(show_dialog=False):
            QMessageBox.warning(
                self,
                "PDF Refresh Failed",
                "The receipt was changed, but the invoice PDF could not be refreshed.",
            )
            return False
        if not self.sync_latest_invoice_request_pdf(int(self.invoice_no.text())):
            QMessageBox.warning(
                self,
                "Request PDF Refresh Failed",
                "The invoice PDF was updated, but the latest request history PDF could not be refreshed.",
            )
        self.remove_moved_private_files()
        self.refresh_invoice_payments()
        self.refresh_invoice_history()
        self.refresh_pending_invoices()
        return True


    def open_whatsapp_chat(self, mobile, message):

        whatsapp_web_url = QUrl(
            f"https://web.whatsapp.com/send?phone={mobile}&text={quote(message)}"
        )
        if sys.platform.startswith("linux"):
            return QDesktopServices.openUrl(whatsapp_web_url)

        whatsapp_url = QUrl(
            f"whatsapp://send?phone={mobile}&text={quote(message)}"
        )
        if QDesktopServices.openUrl(whatsapp_url):
            return True
        return QDesktopServices.openUrl(whatsapp_web_url)


    def sync_latest_invoice_request_pdf(self, invoice_no):

        with sqlite3.connect(DB_PATH) as con:
            row = con.execute(
                """
                SELECT i.pdf_path, i.grand_total,
                       COALESCE(SUM(p.amount), 0), r.id, r.pdf_path
                FROM invoices AS i
                JOIN invoice_requests AS r ON r.invoice_no = i.invoice_no
                LEFT JOIN invoice_payments AS p
                  ON p.invoice_no = i.invoice_no
                 AND p.document_kind = i.document_kind
                WHERE i.invoice_no = ? AND i.document_kind = 'INVOICE'
                GROUP BY i.id, r.id
                ORDER BY r.requested_at DESC, r.id DESC
                LIMIT 1
                """,
                (invoice_no,),
            ).fetchone()
        if not row or not row[0] or not row[4]:
            return True
        source = Path(row[0])
        original_request_pdf = Path(row[4])
        if not source.is_file():
            return False
        invoice_total = money(row[1] or 0)
        paid_total = money(row[2] or 0)
        fully_paid = invoice_total > 0 and paid_total >= invoice_total
        request_folder = original_request_pdf.parent
        payment_update_path = request_folder / (
            f"{original_request_pdf.stem}_PAYMENT_STATUS.pdf"
        )
        final_pdf_path = (
            source.with_name(
                f"Invoice_{source.stem.removeprefix('INV-')}_Paid.pdf"
            )
            if fully_paid else None
        )
        try:
            request_folder.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, payment_update_path)
            if final_pdf_path:
                shutil.copy2(source, final_pdf_path)
        except OSError:
            return False
        with sqlite3.connect(DB_PATH) as con:
            con.execute(
                """
                UPDATE invoice_requests
                SET payment_update_path = ?, final_pdf_path = ?
                WHERE id = ?
                """,
                (
                    str(payment_update_path),
                    str(final_pdf_path or ""),
                    row[3],
                ),
            )
        return True


    def edit_selected_invoice_payment(self):

        row = self.invoice_payments_table.currentRow()
        if row < 0:
            return
        payment_id = self.invoice_payments_table.item(row, 0).data(
            Qt.ItemDataRole.UserRole
        )
        with sqlite3.connect(DB_PATH) as con:
            payment = con.execute(
                "SELECT paid_on, amount, method, reference "
                "FROM invoice_payments WHERE id = ? AND invoice_no = ?",
                (payment_id, int(self.invoice_no.text())),
            ).fetchone()
        if not payment:
            self.refresh_invoice_payments()
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Edit Receipt")
        form = QFormLayout(dialog)
        amount_field = NoWheelDoubleSpinBox()
        amount_field.setRange(0.01, 999999999.99)
        amount_field.setDecimals(2)
        amount_field.setPrefix("₹ ")
        amount_field.setValue(float(payment[1]))
        method_field = NoWheelComboBox()
        method_field.addItems(["Cash", "UPI", "Bank Transfer", "Cheque", "Other"])
        method_field.setCurrentText(payment[2])
        date_field = NoWheelDateEdit(
            QDate.fromString(payment[0], "yyyy-MM-dd")
        )
        date_field.setCalendarPopup(True)
        date_field.setDisplayFormat("dd/MM/yyyy")
        reference_field = QLineEdit(payment[3])
        form.addRow("Amount", amount_field)
        form.addRow("Method", method_field)
        form.addRow("Date", date_field)
        form.addRow("Reference", reference_field)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save
            | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        form.addRow(buttons)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        new_amount = money(amount_field.value())
        invoice_no = int(self.invoice_no.text())
        with sqlite3.connect(DB_PATH) as con:
            invoice_total = con.execute(
                "SELECT grand_total FROM invoices "
                "WHERE invoice_no = ? AND document_kind = 'INVOICE'",
                (invoice_no,),
            ).fetchone()[0]
            others_total = con.execute(
                "SELECT COALESCE(SUM(amount), 0) FROM invoice_payments "
                "WHERE invoice_no = ? AND id != ?",
                (invoice_no, payment_id),
            ).fetchone()[0]
        allowed_amount = money(
            Decimal(str(invoice_total or 0)) - Decimal(str(others_total))
        )
        if new_amount > allowed_amount:
            QMessageBox.warning(
                self,
                "Payment Exceeds Balance",
                f"This receipt cannot exceed ₹{allowed_amount:,.2f}.",
            )
            return

        answer = QMessageBox.question(
            self,
            "Confirm Receipt Edit",
            "Apply these changes to the saved receipt?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        with sqlite3.connect(DB_PATH) as con:
            con.execute(
                """
                UPDATE invoice_payments
                SET paid_on = ?, amount = ?, method = ?, reference = ?
                WHERE id = ? AND invoice_no = ?
                """,
                (
                    date_field.date().toString("yyyy-MM-dd"),
                    float(new_amount),
                    method_field.currentText(),
                    reference_field.text().strip(),
                    payment_id,
                    invoice_no,
                ),
            )
        self.refresh_invoice_after_payment_change()


    def delete_selected_invoice_payment(self):

        row = self.invoice_payments_table.currentRow()
        if row < 0:
            return
        item = self.invoice_payments_table.item(row, 0)
        payment_id = item.data(Qt.ItemDataRole.UserRole)
        answer = QMessageBox.question(
            self,
            "Delete Receipt",
            "Delete this payment receipt? This cannot be undone.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        confirmation = QMessageBox.question(
            self,
            "Confirm Permanent Deletion",
            "Confirm again to permanently delete this receipt.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if confirmation != QMessageBox.StandardButton.Yes:
            return
        with sqlite3.connect(DB_PATH) as con:
            con.execute(
                "DELETE FROM invoice_payments WHERE id = ? AND invoice_no = ?",
                (payment_id, int(self.invoice_no.text())),
            )
        self.refresh_invoice_after_payment_change()


    def record_invoice_payment(self):

        if document_kind_for_selection(self.document_kind.currentIndex()) != "INVOICE":
            return

        if not self.save_invoice(show_dialog=False, show_success=False):
            QMessageBox.warning(
                self,
                "Invoice Not Saved",
                "Save a valid invoice before recording a payment.",
            )
            return

        invoice_no = int(self.invoice_no.text())
        amount = money(self.receipt_amount.value())
        if amount <= 0:
            QMessageBox.warning(
                self, "Invalid Payment", "Enter an amount greater than zero."
            )
            return

        with sqlite3.connect(DB_PATH) as con:
            invoice = con.execute(
                "SELECT grand_total FROM invoices "
                "WHERE invoice_no = ? AND document_kind = 'INVOICE'",
                (invoice_no,),
            ).fetchone()
            paid_total = con.execute(
                """
                SELECT COALESCE(SUM(amount), 0)
                FROM invoice_payments
                WHERE document_kind = 'INVOICE' AND invoice_no = ?
                """,
                (invoice_no,),
            ).fetchone()[0]
            balance_due = money(Decimal(str(invoice[0] or 0)) - Decimal(str(paid_total)))
            if amount > balance_due:
                QMessageBox.warning(
                    self,
                    "Payment Exceeds Balance",
                    f"The balance due is ₹{balance_due:,.2f}.",
                )
                return

            con.execute(
                """
                INSERT INTO invoice_payments (
                    document_kind, invoice_no, paid_on, amount, method, reference
                ) VALUES ('INVOICE', ?, ?, ?, ?, ?)
                """,
                (
                    invoice_no,
                    self.receipt_date.date().toString("yyyy-MM-dd"),
                    float(amount),
                    self.receipt_method.currentText(),
                    self.receipt_reference.text().strip(),
                ),
            )

        if not self.save_invoice(show_dialog=False, show_success=False):
            self.refresh_invoice_payments()
            QMessageBox.warning(
                self,
                "Payment Recorded",
                "The payment was recorded, but the invoice workbook could not be refreshed.",
            )
            return
        if not self.export_pdf(show_dialog=False):
            self.refresh_invoice_payments()
            QMessageBox.warning(
                self,
                "Payment Recorded",
                "The payment and workbook were saved, but the PDF could not be refreshed.",
            )
            return
        if not self.sync_latest_invoice_request_pdf(invoice_no):
            QMessageBox.warning(
                self,
                "Request PDF Refresh Failed",
                "The invoice PDF was updated, but the latest request history PDF could not be refreshed.",
            )

        self.reset_receipt_amount_autofill()
        self.receipt_date.setDate(QDate.currentDate())
        self.receipt_reference.clear()
        self.refresh_invoice_payments()
        self.refresh_invoice_request_history()
        self.remove_moved_private_files()
        QMessageBox.information(
            self,
            "Payment Recorded",
            f"₹{amount:,.2f} received by {self.receipt_method.currentText()}.",
        )


    def document_kind_changed(self, index):

        is_invoice_selection = index == 0
        document_kind = document_kind_for_selection(index)
        if self.active_document_kind is None:
            self.active_document_kind = document_kind
        elif document_kind != self.active_document_kind:
            self.save_document_draft(self.active_document_kind)
            self.active_document_kind = document_kind
            if not self.restore_document_draft(document_kind):
                self.new_invoice()
        self.document_number_label.setText(
            "Invoice No" if is_invoice_selection else "Quotation No"
        )
        self.document_details_box.setTitle(
            "Invoice Details" if is_invoice_selection else "Quotation Details"
        )
        self.save_document_button.setText(
            "Save Invoice" if is_invoice_selection else "Save Quotation"
        )
        self.invoice_heading.setText(
            "TAX INVOICE" if is_invoice_selection else "QUOTATION"
        )
        self.new_document_button.setText(
            "New Invoice" if is_invoice_selection else "New Quotation"
        )
        self.private_invoice_checkbox.setText("Private invoice")
        self.private_invoice_checkbox.setVisible(is_invoice_selection)
        if not is_invoice_selection:
            self.private_invoice_checkbox.setChecked(False)
        self.invoice_payments_box.setVisible(is_invoice_selection)
        self.bank_box.setVisible(is_invoice_selection)
        self.invoice_request_history_box.setVisible(is_invoice_selection)
        self.attach_invoice_checkbox.setVisible(is_invoice_selection)
        self.whatsapp_button.setText(
            "Send WhatsApp" if is_invoice_selection else "Send WhatsApp"
        )
        self.whatsapp_button.setVisible(True)
        self.refresh_invoice_payments()
        self.refresh_invoice_request_history()


    def choose_save_folder(self):

        folder = QFileDialog.getExistingDirectory(
            self,
            "Choose a folder for this document",
            str(self.custom_output_dir or Path.home()),
        )
        if folder:
            self.custom_output_dir = Path(folder)
            self.statusBar().showMessage(
                f"Personal save folder selected: {self.custom_output_dir}"
            )


    def share_document_whatsapp(self):

        document_name = self.document_kind.currentText()
        document_kind = document_kind_for_selection(
            self.document_kind.currentIndex()
        )
        is_invoice = document_kind == "INVOICE"
        invoice_total = self.calculate()[3] if is_invoice else Decimal("0.00")
        paid_total = self.invoice_paid_total(int(self.invoice_no.text())) if is_invoice and self.invoice_no.text().isdigit() else Decimal("0.00")
        balance_due = max(Decimal("0.00"), money(invoice_total - paid_total))
        request_percent = self.payment_percent.value() if is_invoice else 0
        request_amount = min(
            money(invoice_total * Decimal(str(request_percent)) / 100),
            balance_due,
        )
        attach_invoice = is_invoice and self.attach_invoice_checkbox.isChecked()
        mobile = re.sub(r"\D", "", self.customer_mobile.text())
        if len(mobile) == 10:
            mobile = "91" + mobile
        if not mobile:
            QMessageBox.warning(
                self, "Buyer Mobile Required",
                "Enter the buyer's mobile number before sharing the document.",
            )
            return

        try:
            invoice_no = int(self.invoice_no.text())
        except ValueError:
            QMessageBox.warning(
                self, f"{document_name} Number Required",
                f"Enter a valid {document_name.lower()} number before sharing.",
            )
            return
        if not self.save_document_bundle(show_dialog=False):
            QMessageBox.warning(
                self,
                f"{document_name} Not Sent",
                f"Save the {document_name.lower()} and PDF before opening WhatsApp.",
            )
            return
        if is_invoice and self.payment_percent.value() != request_percent:
            self.payment_percent.setValue(request_percent)
            if not self.save_document_bundle(show_dialog=False):
                QMessageBox.warning(
                    self,
                    f"{document_name} Not Sent",
                    "The selected payment request could not be saved to the invoice PDF.",
                )
                return
        with sqlite3.connect(DB_PATH) as con:
            row = con.execute(
                """SELECT pdf_path FROM invoices
                   WHERE invoice_no = ? AND document_kind = ?""",
                (
                    invoice_no,
                    document_kind_for_selection(self.document_kind.currentIndex()),
                ),
            ).fetchone()
        pdf_path = row[0] if row else ""
        if not pdf_path or not Path(pdf_path).is_file():
            QMessageBox.warning(
                self,
                "PDF Missing",
                "Save the document PDF before attaching it in WhatsApp.",
            )
            return
        attachment_paths = (
            [Path(pdf_path).resolve()]
            if not is_invoice or attach_invoice else []
        )
        qr_path = None
        if is_invoice and balance_due > 0 and self.bank_upi_id.text().strip():
            qr_path = self.quotation_payment_qr_path(
                invoice_no, self.invoice_date.date().toPython()
            )
            if not qr_path.is_file():
                QMessageBox.warning(
                    self,
                    "Payment QR Missing",
                    f"The invoice scanner could not be found:\n{qr_path}",
                )
                return
            attachment_paths.append(qr_path.resolve())

        if is_invoice:
            request_time = datetime.now()
            timestamp = request_time.strftime("%Y%m%d_%H%M%S_%f")
            request_folder = Path(pdf_path).parent / "Requests"
            request_folder.mkdir(parents=True, exist_ok=True)
            snapshot_pdf_path = request_folder / (
                f"{Path(pdf_path).stem}_REQUEST_{timestamp}.pdf"
            )
            shutil.copy2(pdf_path, snapshot_pdf_path)
            snapshot_qr_path = None
            if qr_path and qr_path.is_file():
                snapshot_qr_path = request_folder / (
                    f"{Path(pdf_path).stem}_REQUEST_{timestamp}_QR.png"
                )
                shutil.copy2(qr_path, snapshot_qr_path)
            with sqlite3.connect(DB_PATH) as con:
                con.execute(
                    """
                    INSERT INTO invoice_requests (
                        invoice_no, requested_at, request_percent, request_amount,
                        pdf_path, qr_path
                    ) VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        invoice_no,
                        request_time.isoformat(timespec="microseconds"),
                        request_percent,
                        float(request_amount),
                        str(snapshot_pdf_path),
                        str(snapshot_qr_path or ""),
                    ),
                )
            attachment_paths = (
                [snapshot_pdf_path.resolve()] if attach_invoice else []
            )
            if snapshot_qr_path:
                attachment_paths.append(snapshot_qr_path.resolve())
            self.refresh_invoice_request_history()

        pdf_mime_data = QMimeData()
        pdf_mime_data.setUrls(
            [QUrl.fromLocalFile(str(path)) for path in attachment_paths]
        )
        QApplication.clipboard().setMimeData(pdf_mime_data)
        greeting = self.customer_name.text().strip() or "Customer"
        if not is_invoice:
            message = (
                f"Hello {greeting},\n\n"
                "Please find attached our quotation for your review. "
                "We would appreciate your feedback and look forward to the opportunity to work with you. \n\n"
                "Please feel free to reach out if you have any questions or require any changes.\n\n"
                f"Regards,\n{self.seller_name.text()}"
            )
        else:
            message_lines = [
                f"Hello {greeting},",
                self.seller_name.text(),
                "",
                f"Invoice total: INR {invoice_total:.2f}",
                f"Payment request ({request_percent:g}%): INR {request_amount:.2f}",
                f"Received: INR {paid_total:.2f}",
            ]
            if balance_due > 0:
                message_lines.append(f"Balance due: INR {balance_due:.2f}")
            else:
                message_lines.append("Payment complete")
            message = "\n".join(message_lines)
        opened = self.open_whatsapp_chat(mobile, message)
        attachment_names = " and ".join(path.name for path in attachment_paths)
        if opened:
            QMessageBox.information(
                self,
                "Attach Files",
                f"The following file(s) are copied. Paste them into the WhatsApp chat (Ctrl+V):\n\n{attachment_names}",
            )
        else:
            QMessageBox.warning(
                self,
                "WhatsApp Could Not Be Opened",
                f"The following file(s) are copied. Open WhatsApp and paste them into the chat (Ctrl+V):\n{attachment_names}",
            )


    def upload_signature(self):

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select signature image",
            str(Path.home()),
            "PNG images (*.png)",
        )
        if not file_path:
            return

        image_reader = QImageReader(file_path)
        image_format = bytes(image_reader.format()).lower()
        if not image_reader.canRead() or image_format != b"png":
            QMessageBox.warning(
                self,
                "Invalid Signature Image",
                "Select a valid PNG image.",
            )
            return

        try:
            shutil.copy2(file_path, SIGNATURE_PATH)
        except OSError as error:
            QMessageBox.warning(
                self,
                "Signature Could Not Be Saved",
                str(error),
            )
            return

        self.signature_image_path = SIGNATURE_PATH
        self.refresh_signature_display()


    def refresh_signature_display(self):

        if not self.signature_image_path or not self.signature_image_path.is_file():
            self.signature_preview.setPixmap(QPixmap())
            self.signature_preview.setText("No signature")
            self.signature_image_label.setText("No signature uploaded")
            self.upload_signature_button.setText("Upload signature PNG")
            return

        image = QImage(str(self.signature_image_path))
        if image.isNull():
            self.signature_preview.setPixmap(QPixmap())
            self.signature_preview.setText("Preview unavailable")
        else:
            self.signature_preview.setText("")
            self.signature_preview.setPixmap(
                QPixmap.fromImage(image).scaled(
                    self.signature_preview.size(),
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation,
                )
            )
        self.signature_image_label.setText("Signature ready for invoices")
        self.upload_signature_button.setText("Replace signature")


    def add_item_row(self):

        if self.items.rowCount() >= 10:
            QMessageBox.information(
                self,
                "Item Limit",
                "The invoice sheet supports up to 10 item rows."
            )
            return

        row = self.items.rowCount()
        self.items.insertRow(row)
        self.items.blockSignals(True)
        for column in range(self.items.columnCount()):
            item = QTableWidgetItem("")
            item.setBackground(QBrush(QColor("#ffffff")))
            item.setForeground(QBrush(QColor("#1f2937")))
            if column == 5:
                item.setText("PCS")
            if column not in (1, 2, 3, 4, 5):
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.items.setItem(row, column, item)
        self.items.blockSignals(False)
        self.refresh_item_row_controls()
        self.refresh()


    def refresh_item_row_controls(self):

        for row in range(self.items.rowCount()):
            remove_button = QPushButton("Remove")
            remove_button.clicked.connect(
                lambda checked=False, item_row=row: self.remove_item_row(item_row)
            )
            self.items.setCellWidget(row, 7, remove_button)


    def remove_item_row(self, row=None):

        if row is None:
            row = self.items.currentRow()
        if row >= 0:
            self.items.removeRow(row)
            self.refresh_item_row_controls()
            self.refresh()


    # ========================================================
    # NEW INVOICE
    # ========================================================

    def new_invoice(self):

        document_kind = document_kind_for_selection(
            self.document_kind.currentIndex()
        )
        self.delete_document_draft(document_kind)
        self.loaded_invoice_no = None
        self.custom_output_dir = None
        self.reset_receipt_amount_autofill()
        self.private_invoice_checkbox.setChecked(False)

        self.invoice_no.setText(
            str(
                self.next_invoice_number()
            )
        )


        self.invoice_date.setDate(
            QDate.currentDate()
        )


        self.customer_name.clear()
        self.buyer_selector.setCurrentIndex(0)

        self.customer_gstin.clear()

        self.customer_address.clear()
        self.customer_address_2.clear()

        self.customer_mobile.clear()


        self.customer_address.clear()
        self.customer_address_2.clear()
        self.customer_mobile.clear()
        self.po_number.clear()
        self.amount_words.clear()
        self.tax_words.clear()
        self.bundle_count.clear()
        self.through.clear()
        self.items.setRowCount(0)
        self.invoice_type.setCurrentIndex(0)
        self.gst_rate.setValue(5)
        self.payment_percent.setValue(50)


        self.refresh()
        self.refresh_invoice_payments()
        self.refresh_invoice_request_history()
        self.invoice_scroll_area.verticalScrollBar().setValue(0)


        self.statusBar().showMessage(
            f"New {self.document_kind.currentText().lower()} created"
        )


    # ========================================================
    # READ ITEMS
    # ========================================================

    def cell(
        self,
        row,
        column
    ):

        item = self.items.item(
            row,
            column
        )

        if item:

            return item.text().strip()

        return ""


    def decimal(
        self,
        value
    ):

        try:

            return Decimal(
                str(value or "0")
            )

        except:

            return Decimal("0")


    def get_items(self):

        items = []


        for row in range(
            self.items.rowCount()
        ):

            description = self.cell(
                row,
                1
            )


            if not description:

                continue


            quantity = self.decimal(
                self.cell(
                    row,
                    3
                )
            )


            rate = self.decimal(
                self.cell(
                    row,
                    4
                )
            )


            per = self.cell(
                row,
                5
            )

            if not per:

                per = "PCS"


            hsn = self.cell(
                row,
                2
            )


            amount = money(
                quantity * rate
            )


            items.append({

                "description":
                    description,

                "quantity":
                    quantity,

                "rate":
                    rate,

                "per":
                    per,

                "hsn":
                    hsn,

                "amount":
                    amount

            })


        return items


    # ========================================================
    # CALCULATE
    # ========================================================

    def calculate(self):

        items = self.get_items()


        subtotal = money(
            sum(
                (
                    item["amount"]
                    for item in items
                ),
                Decimal("0")
            )
        )


        gst = money(

            subtotal
            *
            Decimal(
                str(
                    self.gst_rate.value()
                )
            )
            /
            Decimal("100")

        )


        total = money(
            subtotal + gst
        )


        return (
            items,
            subtotal,
            gst,
            total
        )


    def split_gst(self, gst):

        cgst = money(gst / 2)
        sgst = money(gst - cgst)

        return cgst, sgst


    # ========================================================
    # REFRESH TOTAL
    # ========================================================

    def refresh(self):

        if not hasattr(
            self,
            "summary"
        ):

            return


        (
            items,
            subtotal,
            gst,
            total
        ) = self.calculate()

        self.items.blockSignals(True)

        for row in range(self.items.rowCount()):
            self.items.item(row, 0).setText(str(row + 1))
            amount = money(
                self.decimal(self.cell(row, 3))
                * self.decimal(self.cell(row, 4))
            ) if self.cell(row, 1) else Decimal("0.00")
            self.items.item(row, 6).setText(f"{amount:.2f}")

        self.items.blockSignals(False)


        if self.invoice_type.currentIndex() == 0:

            tax_text = (
                f"IGST "
                f"{self.gst_rate.value():g}%"
                f": ₹{gst:.2f}"
            )

        else:
            cgst, sgst = self.split_gst(gst)

            tax_text = (

                f"CGST "
                f"{self.gst_rate.value()/2:g}%"
                f": ₹{cgst:.2f}    "

                f"SGST "
                f"{self.gst_rate.value()/2:g}%"
                f": ₹{sgst:.2f}"

            )


        self.summary.setText(

            f"Taxable Value: "
            f"₹{subtotal:.2f}     "

            f"{tax_text}     "

            f"Grand Total: "
            f"₹{total:.2f}"

        )
        self.refresh_payment_request()

        self.amount_words.setText(amount_to_words(total))
        self.tax_words.setText(amount_to_words(gst))

        cgst, sgst = self.split_gst(gst)
        if self.invoice_type.currentIndex() == 0:
            tax_headers = [
                "HSN/SAC", "", "Taxable Value", "Integrated Rate",
                "Integrated Amount", "", "", "Total Tax Amount"
            ]
            tax_values = [
                self.hsn_code.text(), "", f"{subtotal:.2f}",
                f"{self.gst_rate.value():g}%", f"{gst:.2f}", "", "",
                f"{gst:.2f}"
            ]
        else:
            tax_headers = [
                "HSN/SAC", "", "Taxable Value", "Central Rate",
                "Central Amount", "State Rate", "State Amount",
                "Total Tax Amount"
            ]
            tax_values = [
                self.hsn_code.text(), "", f"{subtotal:.2f}",
                f"{self.gst_rate.value() / 2:g}%", f"{cgst:.2f}",
                f"{self.gst_rate.value() / 2:g}%", f"{sgst:.2f}",
                f"{gst:.2f}"
            ]

        self.hsn_table.setHorizontalHeaderLabels(tax_headers)
        for column, value in enumerate(tax_values):
            self.hsn_table.setItem(0, column, QTableWidgetItem(value))
        self.hsn_table.setItem(1, 0, QTableWidgetItem("TOTAL"))
        self.hsn_table.setItem(1, 2, QTableWidgetItem(f"{subtotal:.2f}"))
        for column in (4, 6, 7):
            item = self.hsn_table.item(0, column)
            if item and item.text():
                self.hsn_table.setItem(1, column, QTableWidgetItem(item.text()))


    # ========================================================
    # SAVE XLSX
    # ========================================================

    def create_invoice_workbook(
        self,
        items,
        subtotal,
        gst,
        total,
        invoice_no,
        invoice_date,
        invoice_type,
        document_kind
    ):

        workbook = load_workbook(DATA_DIR / "C GST.xlsx")
        sheet = workbook.active
        paid_total = (
            self.invoice_paid_total(invoice_no)
            if document_kind == "INVOICE" else Decimal("0.00")
        )
        balance_due = max(Decimal("0.00"), money(total - paid_total))
        received_percent = float(paid_total / total * 100) if total else 0
        authorised_signatory = sheet["F50"].value or "Authorised Signatory"
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None:
                    cell.value = None
        sheet.title = "MAIN"
        sheet.sheet_view.showGridLines = False
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.orientation = "portrait"
        sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 1
        sheet.page_setup.scale = None
        sheet.print_area = "A1:H51"
        existing_merges = {str(cell_range) for cell_range in sheet.merged_cells.ranges}
        for merge_range in ("A8:H8", "A9:H9"):
            if merge_range in existing_merges:
                sheet.unmerge_cells(merge_range)
                existing_merges.remove(merge_range)
        for address in ("A8", "A9"):
            current_border = sheet[address].border
            sheet[address].border = Border(
                left=current_border.left,
                right=Side(),
                top=current_border.top,
                bottom=current_border.bottom,
            )
        for merge_range in ("A8:D8", "E8:H8", "A9:D9", "E9:H9"):
            sheet.merge_cells(merge_range)
            existing_merges.add(merge_range)
        contact_edge = Side(style="medium")
        for row in (8, 9):
            cell = sheet.cell(row, 8)
            current_border = cell.border
            cell.border = Border(
                left=current_border.left,
                right=contact_edge,
                top=current_border.top,
                bottom=current_border.bottom,
            )
        for column in range(1, 9):
            cell = sheet.cell(9, column)
            current_border = cell.border
            cell.border = Border(
                left=current_border.left,
                right=current_border.right,
                top=current_border.top,
                bottom=contact_edge,
            )
        if document_kind == "INVOICE":
            for merge_range in (
                "D43:H43", "D44:E44", "F44:H44", "D45:E45", "F45:H45",
                "D46:E46", "F46:H46", "F47:H47", "F48:H49", "F50:H50",
            ):
                if merge_range in existing_merges:
                    sheet.unmerge_cells(merge_range)
                    existing_merges.remove(merge_range)
            for merge_range in (
                "A43:C43", "A44:C44", "A45:C45", "D43:F43",
                "D44:F44", "D45:F45", "D46:F46", "G43:H43",
                "G44:G46", "A47:C47", "A48:E49", "F47:H47",
                "F48:H49", "F50:H50", "A51:H51",
            ):
                sheet.merge_cells(merge_range)
                existing_merges.add(merge_range)
        if document_kind == "QUOTATION":
            for merge_range in (
                "A44:C44", "A45:C45", "D44:E44", "F44:H44",
                "D45:E45", "F45:H45", "D46:E46", "F46:H46",
                "D43:H43", "A47:C47", "A48:E49", "F47:H47",
                "F48:H49", "F50:H50", "A51:H51",
            ):
                if merge_range in existing_merges:
                    sheet.unmerge_cells(merge_range)
                    existing_merges.remove(merge_range)
            for merge_range in (
                "A43:C43", "A44:C44", "A45:C45", "D43:F43",
                "D44:F44", "D45:F45", "D46:F46", "G43:H43",
                "G44:G46",
                "A47:C47", "A48:E49", "F47:H47", "F48:H49",
                "F50:H50", "A51:H51",
            ):
                sheet.merge_cells(merge_range)
                existing_merges.add(merge_range)
        for merge_range in (
            *(f"B{row}:C{row}" for row in range(23, 26)),
            "E12:F12", "G12:H12", "E13:F13", "G13:H13",
            "E14:F14", "G14:H14",
            "A34:H34", "A35:H35", "A41:H41", "A42:H42",
            "A44:C44", "A45:C45", "A47:C47", "A48:E49", "A51:H51",
        ):
            if document_kind == "QUOTATION" and merge_range in (
                "A47:C47", "A48:E49", "A51:H51"
            ):
                continue
            if merge_range not in existing_merges:
                sheet.merge_cells(merge_range)

        detail_row, tax_total_row, eoe_row = 38, 39, 33
        amount_label_row, amount_words_row = 34, 35
        tax_words_row = 41

        website_parts = re.split(
            r"\s+PHONE\s*[:\-]*\s*",
            self.seller_website_phone.text(),
            maxsplit=1,
            flags=re.IGNORECASE,
        )
        website_text = website_parts[0].strip()
        website_phone = website_parts[1].strip() if len(website_parts) > 1 else ""
        email_text = self.seller_email_phone.text().strip()
        email_match = re.search(
            r"\s+((?:\+\s*)?\d[\d\s().-]*\d)\s*$",
            email_text,
        )
        email_phone = email_match.group(1).strip() if email_match else ""
        if email_match:
            email_text = email_text[:email_match.start()].strip()

        values = {
            "A1": self.invoice_heading.text(),
            "A2": self.seller_name.text(),
            "A3": self.seller_business.text(),
            "A4": self.seller_address_1.text(),
            "A5": self.seller_address_2.text(),
            "A6": f"GSTIN :{self.seller_gstin.text()}",
            "A7": self.seller_state.text(),
            "A8": website_text,
            "E8": f"PHONE: {website_phone}" if website_phone else "",
            "A9": email_text,
            "E9": email_phone,
            "A10": "BUYER DETAILS",
            "E10": f"{self.document_kind.currentText().upper()} DETAILS",
            "A11": self.customer_name.text(),
            "A12": "\n".join(
                line for line in (
                    self.customer_address.text().strip(),
                    self.customer_address_2.text().strip(),
                ) if line
            ),
            "E12": f"{self.document_kind.currentText().upper()} NO",
            "G12": invoice_no,
            "E13": "DATE",
            "G13": invoice_date.strftime("%d/%m/%Y"),
            "E14": "P.O NO",
            "G14": self.po_number.text(),
            "A14": f"Mobile: {self.customer_mobile.text()}",
            "A15": f"Gst no: {self.customer_gstin.text()}",
            "A16": "SL.NO",
            "B16": "Description of Goods",
            "D16": "HSN/SAC",
            "E16": "Quantity",
            "F16": "Rate",
            "G16": "Per",
            "H16": "Amount",
            "D43": "BANK DETAILS:",
            "A44": f"NO OF BUNDLE'S : {self.bundle_count.text()}",
            "A45": f"THROUGH : {self.through.text()}",
            "D44": "Bank Name:",
            "F44": self.bank_name.text(),
            "D45": "A/C. No:",
            "F45": self.bank_account.text(),
            "D46": "Branch & IFSC Code:",
            "F46": self.bank_ifsc.text(),
            "A47": "Declaration",
            "F47": self.seller_signatory.text(),
            "A48": self.declaration.toPlainText().strip(),
            "F50": authorised_signatory,
        }
        if document_kind == "INVOICE":
            for address in (
                "F44", "F45", "D44", "D45", "F46", "D46", "F47", "F50",
                "A47", "A48", "A44", "A45",
            ):
                values.pop(address, None)
            values.update({
                "A43": "BUNDLES / THROUGH",
                "A44": f"NO OF BUNDLE'S: {self.bundle_count.text()}",
                "A45": f"THROUGH: {self.through.text()}",
                "D43": "BANK DETAILS",
                "D44": f"Bank Name: {self.bank_name.text()}",
                "D45": f"A/C No: {self.bank_account.text()}",
                "D46": f"IFSC: {self.bank_ifsc.text()}",
                "G43": "PAYMENT SUMMARY",
                "H44": f"Payment received: {received_percent:g}%",
                "H45": f"Received: INR {paid_total:.2f}",
                "H46": f"Balance due: INR {balance_due:.2f}",
                "A47": "Declaration",
                "A48": self.declaration.toPlainText().strip(),
                "F47": self.seller_signatory.text(),
                "F50": "Authorised Signatory",
            })
        else:
            for address in (
                "A43", "A44", "A45", "D43", "D44", "F44", "D45",
                "F45", "D46", "F46", "G43", "H44", "H45", "H46",
            ):
                values.pop(address, None)
            for row in range(43, 47):
                sheet.row_dimensions[row].hidden = True
            values.update({
                "A47": "Declaration",
                "A48": self.declaration.toPlainText().strip(),
                "F47": self.seller_signatory.text(),
                "F50": "Authorised Signatory",
            })
        for address, value in values.items():
            sheet[address] = value
        if document_kind == "INVOICE":
            for cell_range in (
                "A43:C43", "D43:F43", "G43:H43",
                "A44:C44", "A45:C45", "D44:F44", "D45:F45",
                "D46:F46", "H44:H46", "A47:C47", "A48:E49",
            ):
                for row_cells in sheet[cell_range]:
                    for cell in row_cells:
                        cell.alignment = Alignment(
                            horizontal="left", vertical="center", wrap_text=True
                        )
            for header in ("A43", "D43", "G43"):
                sheet[header].font = Font(bold=True)
                sheet[header].alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )
            for address in ("F47", "F50"):
                sheet[address].alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )
            for row in range(44, 47):
                sheet.row_dimensions[row].height = 30
            sheet.column_dimensions["G"].width = 13
            sheet.column_dimensions["H"].width = 30
            sheet.row_dimensions[47].height = 20
            sheet.row_dimensions[48].height = 27
            sheet.row_dimensions[49].height = 27
            sheet.row_dimensions[51].height = 21
            footer_side = Side(style="thin")
            footer_border = Border(
                left=footer_side,
                right=footer_side,
                top=footer_side,
                bottom=footer_side,
            )
            signature_area_border = Border(
                left=Side(),
                right=Side(),
                top=Side(),
                bottom=Side(),
            )
            for address in (
                "A43", "D43", "G43", "A44", "A45", "D44", "D45", "D46",
                "G44", "H44", "H45", "H46", "A47", "A48",
            ):
                sheet[address].border = footer_border
            sheet["F48"].border = signature_area_border
            for cell in sheet["A47:C47"][0]:
                cell.border = Border(
                    left=cell.border.left,
                    right=cell.border.right,
                    top=cell.border.top,
                    bottom=Side(),
                )
            for row_cells in sheet["A48:E49"]:
                for cell in row_cells:
                    cell.border = Border(
                        left=cell.border.left,
                        right=Side(),
                        top=Side(),
                        bottom=Side(),
                    )
            for row in range(47, 51):
                cell = sheet.cell(row, 6)
                cell.border = Border(
                    left=Side(),
                    right=cell.border.right,
                    top=cell.border.top,
                    bottom=cell.border.bottom,
                )
            for row in range(47, 51):
                cell = sheet.cell(row, 8)
                cell.border = Border(
                    left=cell.border.left,
                    right=footer_side,
                    top=cell.border.top,
                    bottom=cell.border.bottom,
                )
            for column in range(1, 6):
                cell = sheet.cell(50, column)
                cell.border = Border(
                    left=cell.border.left,
                    right=cell.border.right,
                    top=cell.border.top,
                    bottom=footer_side,
                )
            for row in range(47, 51):
                for column in range(4, 7):
                    cell = sheet.cell(row, column)
                    cell.border = Border(
                        left=Side(),
                        right=Side(),
                        top=cell.border.top,
                        bottom=cell.border.bottom,
                    )
            for column in range(1, 9):
                sheet.cell(51, column).border = Border(
                    left=Side(), right=Side(), top=Side(), bottom=Side()
                )
            for row in range(43, 47):
                for column in range(1, 9):
                    sheet.cell(row, column).border = footer_border
            for column in range(1, 4):
                sheet.cell(46, column).border = Border(
                    left=Side(),
                    right=Side(),
                    top=Side(),
                    bottom=Side(),
                )
            sheet["A2"].alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
        sheet["A12"].alignment = Alignment(
            horizontal="left", vertical="center", wrap_text=True
        )
        if self.customer_address_2.text().strip():
            sheet.row_dimensions[12].height = 30
        sheet[f"A{amount_label_row}"] = "Amount Chargeable (in words) :"
        sheet[f"A{amount_words_row}"] = amount_to_words(total)
        jurisdiction_cell = "A51"
        sheet[jurisdiction_cell].alignment = Alignment(
            horizontal="left", vertical="center", wrap_text=False
        )
        sheet.row_dimensions[51].hidden = False
        sheet[jurisdiction_cell] = self.jurisdiction.text().strip()

        for index, item in enumerate(items):
            row = 18 + index
            sheet.cell(row, 1, index + 1)
            sheet.cell(row, 2, item["description"])
            sheet.cell(row, 4, item["hsn"])
            sheet.cell(row, 5, float(item["quantity"]))
            sheet.cell(row, 6, float(item["rate"]))
            sheet.cell(row, 7, item["per"])
            sheet.cell(row, 8, float(item["amount"]))
            wrapped_lines = sum(
                max(1, len(textwrap.wrap(line, width=52, break_long_words=True)))
                for line in item["description"].splitlines()
            ) or 1
            sheet.row_dimensions[row].height = max(22.5, wrapped_lines * 18)
        for row in range(18 + len(items), 28):
            sheet.row_dimensions[row].height = 6

        sheet["H28"] = float(subtotal)
        tax_rate = self.gst_rate.value()
        cgst, sgst = self.split_gst(gst)
        if invoice_type == "IGST":
            sheet["B29"] = "IGST"
            sheet["H29"] = float(gst)
            sheet["B32"] = "TOTAL"
            sheet["H32"] = float(total)
            sheet["A36"] = "HSN/SAC"
            sheet["C36"] = "Taxable value"
            sheet["D36"] = "Integrated Tax"
            sheet["H36"] = "Total Tax Amount"
            sheet["D37"] = "Rate"
            sheet["E37"] = "Amount"
            sheet.cell(detail_row, 1, self.hsn_code.text())
            sheet.cell(detail_row, 3, float(subtotal))
            sheet.cell(detail_row, 4, tax_rate / 100)
            sheet.cell(detail_row, 5, float(gst))
            sheet.cell(detail_row, 8, float(gst))
            sheet.cell(tax_total_row, 1, "TOTAL")
            sheet.cell(tax_total_row, 3, float(subtotal))
            sheet.cell(tax_total_row, 5, float(gst))
            sheet.cell(tax_total_row, 8, float(gst))
            sheet[f"A{tax_words_row}"] = "Tax Amount (in words) :"
            sheet[f"A{tax_words_row + 1}"] = amount_to_words(gst)
        else:
            sheet["B29"] = "CGST"
            sheet["H29"] = float(cgst)
            sheet["B30"] = "SGST"
            sheet["H30"] = float(sgst)
            sheet["B32"] = "TOTAL"
            sheet["H32"] = float(total)
            sheet["A36"] = "HSN/SAC"
            sheet["C36"] = "Taxable value"
            sheet["D36"] = "Central Tax"
            sheet["F36"] = "State Tax"
            sheet["H36"] = "Total Tax Amount"
            sheet["D37"] = "Rate"
            sheet["E37"] = "Amount"
            sheet["F37"] = "Rate"
            sheet["G37"] = "Amount"
            sheet.cell(detail_row, 1, self.hsn_code.text())
            sheet.cell(detail_row, 3, float(subtotal))
            sheet.cell(detail_row, 4, tax_rate / 200)
            sheet.cell(detail_row, 5, float(cgst))
            sheet.cell(detail_row, 6, tax_rate / 200)
            sheet.cell(detail_row, 7, float(sgst))
            sheet.cell(detail_row, 8, float(gst))
            sheet.cell(tax_total_row, 1, "TOTAL")
            sheet.cell(tax_total_row, 3, float(subtotal))
            sheet.cell(tax_total_row, 5, float(cgst))
            sheet.cell(tax_total_row, 7, float(sgst))
            sheet.cell(tax_total_row, 8, float(gst))
            sheet[f"A{tax_words_row}"] = "Tax Amount (in words) :"
            sheet[f"A{tax_words_row + 1}"] = amount_to_words(gst)

        sheet[f"D{eoe_row}"] = "E. & O.E"
        sheet["A35"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        sheet["A42"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        sheet["A48"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        sheet.row_dimensions[35].height = 27
        sheet.row_dimensions[42].height = 27
        sheet.row_dimensions[48].height = 27
        sheet.row_dimensions[49].height = 27
        sheet.row_dimensions[50].height = 45

        for row in range(18, 28):
            for column in (1, 5, 6, 7, 8):
                sheet.cell(row, column).alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )
            sheet.cell(row, 2).alignment = Alignment(
                horizontal="left", vertical="center", wrap_text=True
            )
        for address in (
            "G12", "G13", "G14", "F47", "F50",
        ):
            sheet[address].alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
        for address in ("E12", "E13", "E14"):
            sheet[address].alignment = Alignment(
                horizontal="left", vertical="center", wrap_text=True
            )
        for address in ("A8", "A9"):
            sheet[address].alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
        for address in ("E8", "E9"):
            sheet[address].alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
        for row in (8, 9):
            sheet.cell(row, 5).font = copy(sheet.cell(row, 1).font)
            sheet.cell(row, 4).border = Border(
                left=sheet.cell(row, 4).border.left,
                right=Side(),
                top=sheet.cell(row, 4).border.top,
                bottom=sheet.cell(row, 4).border.bottom,
            )
            sheet.cell(row, 5).border = Border(
                left=Side(),
                right=sheet.cell(row, 5).border.right,
                top=sheet.cell(row, 5).border.top,
                bottom=sheet.cell(row, 5).border.bottom,
            )
        for address in ("F44", "F45", "F46"):
            sheet[address].alignment = Alignment(
                horizontal="left", vertical="center", wrap_text=True
            )
        thin_side = Side(style="thin")
        bank_cell_border = Border(
            left=thin_side,
            right=thin_side,
            top=thin_side,
            bottom=thin_side,
        )
        medium_side = Side(style="medium")

        def set_cell_border(cell, left=None, right=None, top=None, bottom=None):
            current = cell.border
            cell.border = Border(
                left=left or current.left,
                right=right or current.right,
                top=top or current.top,
                bottom=bottom or current.bottom,
            )

        for column in range(1, 9):
            set_cell_border(sheet.cell(10, column), top=medium_side)
        for row in range(10, 16):
            set_cell_border(sheet.cell(row, 1), left=medium_side)
            set_cell_border(sheet.cell(row, 4), right=medium_side)
            set_cell_border(sheet.cell(row, 5), left=medium_side)
            set_cell_border(sheet.cell(row, 8), right=medium_side)

        for row in range(12, 15):
            for column in range(5, 9):
                set_cell_border(
                    sheet.cell(row, column),
                    left=thin_side if column in (5, 7) else None,
                    right=thin_side if column in (6, 8) else None,
                    top=thin_side,
                    bottom=thin_side,
                )
        for row in range(44, 47):
            for column in ("D", "F"):
                sheet[f"{column}{row}"].border = bank_cell_border

        if self.signature_image_path:
            signature = SpreadsheetImage(str(self.signature_image_path))
            signature_columns = ("F", "G", "H")
            signature_width = 280
            signature_height = 80
            scale = min(
                signature_width / signature.width,
                signature_height / signature.height,
                1,
            )
            signature.width = max(1, int(signature.width * scale))
            signature.height = max(1, int(signature.height * scale))
            signatory_width = sum(
                int((sheet.column_dimensions[column].width or 8.43) * 7 + 5)
                for column in signature_columns
            )
            horizontal_offset = max(0, (signatory_width - signature.width) // 2)
            start_column = 5
            while horizontal_offset >= int(
                (sheet.column_dimensions[chr(ord("A") + start_column)].width or 8.43) * 7 + 5
            ):
                horizontal_offset -= int(
                    (sheet.column_dimensions[chr(ord("A") + start_column)].width or 8.43) * 7 + 5
                )
                start_column += 1
            signature.anchor = OneCellAnchor(
                _from=AnchorMarker(
                    col=start_column,
                    row=47,
                    colOff=pixels_to_EMU(horizontal_offset),
                    rowOff=pixels_to_EMU(5),
                ),
                ext=XDRPositiveSize2D(
                    cx=pixels_to_EMU(signature.width),
                    cy=pixels_to_EMU(signature.height),
                ),
            )
            sheet.add_image(signature)

        if (
            document_kind == "INVOICE"
            and self.bank_upi_id.text().strip()
        ):
            payment_uri = self.payment_uri(invoice_no)
            qr_path = self.quotation_payment_qr_path(
                invoice_no, invoice_date, total
            )
            qr_path.parent.mkdir(parents=True, exist_ok=True)
            qrcode.make(payment_uri).save(qr_path)
            payment_qr = SpreadsheetImage(str(qr_path))
            qr_cell_width = int(
                (sheet.column_dimensions["G"].width or 8.43) * 7 + 5
            )
            qr_cell_height = sum(
                (sheet.row_dimensions[row].height or 15) * 96 / 72
                for row in range(44, 47)
            )
            qr_max_size = min(qr_cell_width - 8, int(qr_cell_height - 8))
            scale = min(
                qr_max_size / payment_qr.width,
                qr_max_size / payment_qr.height,
                1,
            )
            payment_qr.width = max(1, int(payment_qr.width * scale))
            payment_qr.height = max(1, int(payment_qr.height * scale))
            payment_qr.anchor = OneCellAnchor(
                _from=AnchorMarker(
                    col=6,
                    row=43,
                    colOff=pixels_to_EMU(
                        max(0, (qr_cell_width - payment_qr.width) // 2)
                    ),
                    rowOff=pixels_to_EMU(
                        max(0, int((qr_cell_height - payment_qr.height) / 2))
                    ),
                ),
                ext=XDRPositiveSize2D(
                    cx=pixels_to_EMU(payment_qr.width),
                    cy=pixels_to_EMU(payment_qr.height),
                ),
            )
            sheet.add_image(payment_qr)

        uniform_border_side = Side(style="thin", color="000000")
        final_row = 55 if document_kind == "QUOTATION" else 51
        for row in sheet.iter_rows(min_row=1, max_row=final_row, min_col=1, max_col=8):
            for cell in row:
                current = cell.border
                cell.border = Border(
                    left=uniform_border_side if current.left.style else Side(),
                    right=uniform_border_side if current.right.style else Side(),
                    top=uniform_border_side if current.top.style else Side(),
                    bottom=uniform_border_side if current.bottom.style else Side(),
                )
        for row in (28, 29, 30, 31, 32, 37, 38, 39):
            for column in range(1, 9):
                sheet.cell(row, column).alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )
        return workbook

    def remove_moved_private_files(self):

        paths = self.pending_private_move_paths
        self.pending_private_move_paths = []
        for path in paths:
            try:
                Path(path).unlink(missing_ok=True)
            except OSError as error:
                QMessageBox.warning(
                    self,
                    "Private File Cleanup Failed",
                    f"The new document was saved, but this old private file could not be removed:\n{path}\n\n{error}",
                )


    def save_document_bundle(self, show_dialog=True):

        if not self.save_invoice(
            show_dialog=show_dialog,
            show_success=False,
        ):
            if self.private_move_cancelled:
                self.private_move_cancelled = False
            return False

        if not self.export_pdf(show_dialog=False):
            QMessageBox.warning(
                self,
                "PDF Export Failed",
                "The XLSX was saved, but the PDF could not be created. Check that LibreOffice is installed.",
            )
            return False

        self.remove_moved_private_files()

        invoice_no = int(self.invoice_no.text())
        document_kind = document_kind_for_selection(
            self.document_kind.currentIndex()
        )
        with sqlite3.connect(DB_PATH) as con:
            xlsx_path, pdf_path = con.execute(
                """SELECT xlsx_path, pdf_path FROM invoices
                   WHERE invoice_no = ? AND document_kind = ?""",
                (invoice_no, document_kind),
            ).fetchone()
        downloads_location = QStandardPaths.writableLocation(
            QStandardPaths.StandardLocation.DownloadLocation
        )
        downloads_dir = Path(downloads_location) if downloads_location else Path.home() / "Downloads"
        downloads_dir.mkdir(parents=True, exist_ok=True)
        download_copy = downloads_dir / Path(pdf_path).name
        try:
            if Path(pdf_path).resolve() != download_copy.resolve():
                shutil.copy2(pdf_path, download_copy)
        except OSError as error:
            if show_dialog:
                QMessageBox.warning(
                    self,
                    "Downloads Copy Failed",
                    f"The PDF was saved to its normal folder, but could not be copied to Downloads:\n{error}",
                )
            return True
        if show_dialog:
            QMessageBox.information(
                self,
                "Document Saved",
                f"XLSX saved to:\n{xlsx_path}\n\nPDF saved to:\n{pdf_path}"
                f"\n\nPDF copy in Downloads:\n{download_copy}",
            )
        return True


    def save_invoice(self, show_dialog=True, show_success=None):

        try:
            if show_success is None:
                show_success = show_dialog
            self.pending_private_move_paths = []
            self.private_move_cancelled = False
            invoice_no = int(self.invoice_no.text())
            invoice_date = self.invoice_date.date().toPython()
            items, subtotal, gst, total = self.calculate()
            if not items:
                if show_dialog:
                    QMessageBox.warning(self, "No Items", "Please enter at least one item.")
                return False

            mobile = self.customer_mobile.text().strip()
            if not normalize_indian_mobile(mobile):
                if show_dialog:
                    QMessageBox.warning(
                        self,
                        "Invalid Buyer Mobile",
                        "Enter a valid Indian mobile number with 10 digits, optionally prefixed by +91.",
                    )
                self.customer_mobile.setFocus()
                return False

            self.save_seller_profile()
            invoice_type = "IGST" if self.invoice_type.currentIndex() == 0 else "CGST + SGST"
            document_kind = document_kind_for_selection(
                self.document_kind.currentIndex()
            )
            scanner_path = None
            document_root = self.invoice_storage_root(
                document_kind, invoice_no, total
            )
            document_folder = document_folder_for_selection(
                self.document_kind.currentIndex()
            )
            year_folder = document_root / document_folder / "XLSX" / str(invoice_date.year)
            year_folder.mkdir(parents=True, exist_ok=True)
            output = year_folder / invoice_filename(
                invoice_no, self.customer_name.text(), ".xlsx", document_kind
            )

            with sqlite3.connect(DB_PATH) as con:
                existing_invoice = con.execute(
                    "SELECT xlsx_path, pdf_path, customer_name, document_kind "
                    "FROM invoices WHERE invoice_no = ? AND document_kind = ?",
                    (invoice_no, document_kind),
                ).fetchone()
            paid_total = (
                self.invoice_paid_total(invoice_no)
                if document_kind == "INVOICE" else Decimal("0.00")
            )
            if paid_total > money(total):
                if show_dialog:
                    QMessageBox.warning(
                        self,
                        "Invoice Total Below Amount Received",
                        f"The recorded payments total ₹{paid_total:,.2f}, which is "
                        f"more than the new invoice total of ₹{money(total):,.2f}.",
                    )
                return False
            if existing_invoice and existing_invoice[0]:
                previous_xlsx = Path(existing_invoice[0])
                was_private = previous_xlsx.resolve().is_relative_to(
                    PRIVATE_INVOICE_DIR.resolve()
                )
                will_be_private = output.resolve().is_relative_to(
                    PRIVATE_INVOICE_DIR.resolve()
                )
                if was_private and not will_be_private:
                    answer = QMessageBox.question(
                        self,
                        "Move Private Document?",
                        f"This document is already saved in Private/{document_folder}. "
                        f"Move it to {document_folder}? The old private XLSX and PDF "
                        "will be removed after the new PDF is saved.",
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                        QMessageBox.StandardButton.No,
                    )
                    if answer != QMessageBox.StandardButton.Yes:
                        self.private_move_cancelled = True
                        return False
                    self.pending_private_move_paths = [
                        existing_invoice[0],
                        existing_invoice[1],
                    ]
                    if existing_invoice[3] in ("INVOICE", "QUOTATION"):
                        private_qr_path = (
                            PRIVATE_INVOICE_DIR
                            / document_folder
                            / "QR"
                            / str(invoice_date.year)
                            / invoice_filename(
                                invoice_no,
                                existing_invoice[2],
                                "_payment.png",
                                existing_invoice[3],
                            )
                        )
                        self.pending_private_move_paths.append(private_qr_path)
            if existing_invoice and existing_invoice[0]:
                previous_xlsx = Path(existing_invoice[0])
                previous_pdf = Path(existing_invoice[1]) if existing_invoice[1] else None
                target_pdf = (
                    document_root / document_folder / "PDF" / str(invoice_date.year)
                    / f"{output.stem}.pdf"
                )
                cleanup_paths = (previous_xlsx, previous_pdf)
                for old_path in cleanup_paths:
                    if (
                        old_path
                        and old_path.resolve() not in {
                            output.resolve(), target_pdf.resolve()
                        }
                        and old_path not in self.pending_private_move_paths
                    ):
                        self.pending_private_move_paths.append(old_path)
                if previous_xlsx.parent.parent.name == "XLSX":
                    previous_root = previous_xlsx.parent.parent.parent
                    old_qr = (
                        previous_root / "QR" / str(invoice_date.year)
                        / invoice_filename(
                            invoice_no,
                            existing_invoice[2],
                            "_payment.png",
                            document_kind,
                        )
                    )
                    new_qr = (
                        document_root / document_folder / "QR"
                        / str(invoice_date.year)
                        / invoice_filename(
                            invoice_no,
                            self.customer_name.text(),
                            "_payment.png",
                            document_kind,
                        )
                    )
                    if old_qr.resolve() != new_qr.resolve():
                        self.pending_private_move_paths.append(old_qr)
            if output.exists() and not existing_invoice:
                if show_dialog:
                    QMessageBox.warning(
                        self, "Already Exists", f"Invoice {invoice_no} already exists."
                    )
                return False

            workbook = self.create_invoice_workbook(
                items, subtotal, gst, total, invoice_no, invoice_date,
                invoice_type, document_kind,
            )
            workbook.save(output)

            with sqlite3.connect(DB_PATH) as con:
                values = (
                    invoice_date.isoformat(),
                    invoice_type,
                    self.customer_name.text(),
                    float(total),
                    str(output),
                    document_kind,
                    scanner_path,
                    self.bank_upi_id.text().strip(),
                    invoice_no,
                )
                if existing_invoice:
                    con.execute(
                        """
                        UPDATE invoices
                        SET invoice_date = ?, invoice_type = ?, customer_name = ?,
                            grand_total = ?, xlsx_path = ?, document_kind = ?,
                            qr_scanner_path = ?, payment_address = ?
                        WHERE invoice_no = ? AND document_kind = ?
                        """,
                        (*values, document_kind),
                    )
                else:
                    con.execute(
                        """
                        INSERT INTO invoices (
                            invoice_date, invoice_type, customer_name, grand_total,
                            xlsx_path, document_kind, qr_scanner_path, payment_address,
                            invoice_no
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        values,
                    )

            self.loaded_invoice_no = invoice_no
            self.save_document_draft(document_kind)
            self.refresh_invoice_history()
            self.refresh_pending_invoices()
            self.refresh_quotation_register()
            self.refresh_invoice_payments()
            self.statusBar().showMessage(f"{document_kind.title()} saved: {output}")
            if show_dialog and show_success:
                QMessageBox.information(
                    self,
                    f"{document_kind.title()} Saved",
                    f"{document_kind.title()} {invoice_no} saved successfully.\n\n{output}",
                )
            return True

        except sqlite3.IntegrityError:
            if show_dialog:
                QMessageBox.warning(
                    self, "Already Saved", "This invoice number already exists."
                )
            return False
        except Exception as error:
            if show_dialog:
                QMessageBox.critical(self, "Error", str(error))
            return False


    # ========================================================
    # PDF
    # ========================================================

    def export_pdf(self, show_dialog=True):

        try:

            invoice_no = int(
                self.invoice_no.text()
            )


            invoice_date = (
                self.invoice_date
                .date()
                .toPython()
            )


            xlsx_file, saved_buyer_name = self.saved_invoice_file(invoice_no)
            document_kind = document_kind_for_selection(
                self.document_kind.currentIndex()
            )
            document_folder = document_folder_for_selection(
                self.document_kind.currentIndex()
            )
            default_year_folder = (
                (
                    self.custom_output_dir
                    or (
                        PRIVATE_INVOICE_DIR
                        if self.private_invoice_checkbox.isChecked()
                        else INVOICE_DIR
                    )
                )
                / document_folder / "XLSX" / str(invoice_date.year)
            )
            buyer_name = saved_buyer_name or self.customer_name.text()
            if not xlsx_file or not xlsx_file.exists():
                preferred_xlsx = default_year_folder / invoice_filename(
                    invoice_no,
                    buyer_name,
                    ".xlsx",
                    document_kind,
                )
                xlsx_file = preferred_xlsx


            # Save XLSX first
            if not xlsx_file.exists():
                if not self.save_invoice(show_dialog=show_dialog):
                    return False
                saved_xlsx, saved_buyer_name = self.saved_invoice_file(invoice_no)
                if saved_xlsx:
                    xlsx_file = saved_xlsx
                    buyer_name = saved_buyer_name or buyer_name


            if not xlsx_file.exists():
                return False


            if len(xlsx_file.parents) > 2 and xlsx_file.parent.parent.name == "XLSX":
                document_root = xlsx_file.parent.parent.parent
            else:
                document_root = INVOICE_DIR / document_folder
            pdf_folder = document_root / "PDF" / str(invoice_date.year)
            pdf_folder.mkdir(parents=True, exist_ok=True)
            pdf_file = pdf_folder / f"{xlsx_file.stem}.pdf"


            if sys.platform == "win32":
                generated_pdf = pdf_file
                powershell = shutil.which("powershell.exe")
                if not powershell:
                    raise RuntimeError(
                        "Windows PowerShell was not found. The Excel invoice "
                        f"is saved at:\n{xlsx_file}"
                    )

                workbook_path = str(xlsx_file.resolve()).replace("'", "''")
                pdf_path = str(pdf_file.resolve()).replace("'", "''")
                script = (
                    "$excel = $null; $book = $null; "
                    "try { "
                    "$excel = New-Object -ComObject Excel.Application; "
                    "$excel.Visible = $false; "
                    f"$book = $excel.Workbooks.Open('{workbook_path}'); "
                    f"$book.ExportAsFixedFormat(0, '{pdf_path}'); "
                    "} finally { "
                    "if ($book) { $book.Close($false) }; "
                    "if ($excel) { $excel.Quit() } "
                    "}"
                )
                result = subprocess.run(
                    [powershell, "-NoProfile", "-NonInteractive", "-Command", script],
                    capture_output=True,
                    text=True
                )
                if result.returncode:
                    raise RuntimeError(
                        "Microsoft Excel could not export the invoice to PDF. "
                        "Confirm Excel is installed and activated.\n\n"
                        f"{result.stderr or result.stdout}"
                    )
            else:
                libreoffice = shutil.which("libreoffice") or shutil.which("soffice")
                if not libreoffice:
                    if show_dialog:
                        QMessageBox.warning(
                            self,
                            "PDF Export Unavailable",
                            "LibreOffice is required to export PDF on Ubuntu. "
                            "The XLSX invoice is already saved:\n"
                            f"{xlsx_file}"
                        )
                    return False

                with tempfile.TemporaryDirectory(dir=DATA_DIR) as temp_dir:
                    result = subprocess.run(
                        [
                            libreoffice,
                            "--headless",
                            "--convert-to", "pdf",
                            "--outdir", temp_dir,
                            str(xlsx_file)
                        ],
                        capture_output=True,
                        text=True
                    )
                    temporary_pdf = Path(temp_dir) / f"{xlsx_file.stem}.pdf"
                    if temporary_pdf.exists():
                        shutil.copy2(temporary_pdf, pdf_file)
                generated_pdf = pdf_file


            if not generated_pdf.exists():

                raise RuntimeError(

                    "PDF was not created.\n\n"
                    +
                    result.stderr

                )


            # Database
            with sqlite3.connect(
                DB_PATH
            ) as con:

                con.execute(

                    """
                    UPDATE invoices

                    SET pdf_path = ?

                    WHERE invoice_no = ? AND document_kind = ?
                    """,

                    (
                        str(
                            generated_pdf
                        ),
                            invoice_no,
                            document_kind,
                    )

                )


            self.statusBar().showMessage(

                f"PDF saved: "
                f"{generated_pdf.name}"

            )

            if show_dialog:
                QMessageBox.information(
                    self,
                    "PDF Saved",
                    f"PDF successfully saved to:\n\n{generated_pdf}"
                )
            return True


        except Exception as error:
            if show_dialog:
                QMessageBox.critical(self, "PDF Error", str(error))
            return False


# ============================================================
# START
# ============================================================

if __name__ == "__main__":

    app = QApplication(
        sys.argv
    )

    app.setApplicationName(
        "Invoice Manager"
    )


    window = InvoiceApp()

    window.show()


    sys.exit(
        app.exec()
    )