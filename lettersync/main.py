from pathlib import Path
import os
import sys
import sqlite3
import traceback
from datetime import datetime, date, timedelta

from PySide6.QtCore import (
    Qt,
    QObject,
    QThread,
    Signal,
    QTime,
    QDate,
    QUrl,
)
from PySide6.QtGui import (
    QDesktopServices,
    QPainter,
    QPen,
    QBrush,
    QColor,
)
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QFormLayout,
    QGridLayout,
    QLabel,
    QPushButton,
    QLineEdit,
    QComboBox,
    QDateEdit,
    QTimeEdit,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QMessageBox,
    QProgressDialog,
    QGroupBox,
    QHeaderView,
    QTextEdit,
    QFrame,
    QGraphicsDropShadowEffect,
)

from app.database import get_connection, init_db
from app.generator import render_template, create_pdf_from_docx


# ============================================================
# PATH

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent

OUTPUT_DIR = BASE_DIR / "output"
WORD_DIR = OUTPUT_DIR / "word"
PDF_DIR = OUTPUT_DIR / "pdf"
TEMPLATE_DIR = BASE_DIR / "templates"
WORD_DIR.mkdir(parents=True, exist_ok=True)
PDF_DIR.mkdir(parents=True, exist_ok=True)


# ============================================================
# DATA INDONESIA
# ============================================================

BULAN_ID = [
    "Januari",
    "Februari",
    "Maret",
    "April",
    "Mei",
    "Juni",
    "Juli",
    "Agustus",
    "September",
    "Oktober",
    "November",
    "Desember",
]

HARI_ID = [
    "Senin",
    "Selasa",
    "Rabu",
    "Kamis",
    "Jumat",
    "Sabtu",
    "Minggu",
]


# ============================================================
# DATABASE
# ============================================================

def ensure_archive_columns():
    """
    Memastikan tabel surat mempunyai kolom tambahan
    yang diperlukan oleh LetterSync.
    """

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("PRAGMA table_info(surat)")
    columns = {row[1] for row in cur.fetchall()}

    if "pdf_path" not in columns:
        cur.execute(
            "ALTER TABLE surat ADD COLUMN pdf_path TEXT"
        )

    if "created_at" not in columns:
        cur.execute(
            "ALTER TABLE surat ADD COLUMN created_at TEXT"
        )

    conn.commit()
    conn.close()


def ensure_lembur_table():
    """
    Membuat tabel lembur apabila belum ada.

    Juga melakukan migrasi otomatis untuk database lama,
    terutama kolom tanggal_surat.
    """

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS lembur (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pegawai_id INTEGER,
            tanggal TEXT NOT NULL,
            tanggal_surat TEXT,
            jam_mulai TEXT,
            jam_selesai TEXT,
            rentang_jam TEXT,
            total_jam TEXT,
            laporan_lembur TEXT DEFAULT '',
            status TEXT DEFAULT 'Draft',
            file_path TEXT,
            pdf_path TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(pegawai_id) REFERENCES pegawai(id)
        )
    """)

    # --------------------------------------------------------
    # MIGRASI DATABASE LAMA
    # --------------------------------------------------------

    cur.execute("PRAGMA table_info(lembur)")
    columns = {row[1] for row in cur.fetchall()}

    if "tanggal_surat" not in columns:
        cur.execute(
            "ALTER TABLE lembur ADD COLUMN tanggal_surat TEXT"
        )

    if "pdf_path" not in columns:
        cur.execute(
            "ALTER TABLE lembur ADD COLUMN pdf_path TEXT"
        )

    if "created_at" not in columns:
        cur.execute(
            "ALTER TABLE lembur ADD COLUMN created_at TEXT"
        )

    if "status" not in columns:
        cur.execute(
            "ALTER TABLE lembur ADD COLUMN status TEXT DEFAULT 'Draft'"
        )

    if "file_path" not in columns:
        cur.execute(
            "ALTER TABLE lembur ADD COLUMN file_path TEXT"
        )

    if "laporan_lembur" not in columns:
        cur.execute(
            "ALTER TABLE lembur ADD COLUMN laporan_lembur TEXT DEFAULT ''"
        )

    conn.commit()
    conn.close()


# ============================================================
# FORMAT TANGGAL
# ============================================================

def parse_date(value):
    if isinstance(value, date):
        return value

    if not value:
        return date.today()

    value = str(value).strip()

    for fmt in (
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%Y-%m-%d",
    ):
        try:
            return datetime.strptime(
                value,
                fmt
            ).date()
        except ValueError:
            pass

    raise ValueError(
        f"Format tanggal tidak dikenali: {value}"
    )


def format_date_id(value):
    d = parse_date(value)

    return (
        f"{d.day} "
        f"{BULAN_ID[d.month - 1]} "
        f"{d.year}"
    )


def format_day_id(value):
    d = parse_date(value)
    return HARI_ID[d.weekday()]


def format_range_day(start, end):
    start = parse_date(start)
    end = parse_date(end)

    if start == end:
        return HARI_ID[start.weekday()]

    return (
        f"{HARI_ID[start.weekday()]} "
        f"s.d {HARI_ID[end.weekday()]}"
    )


def format_range_date(start, end):
    start = parse_date(start)
    end = parse_date(end)

    if start == end:
        return format_date_id(start)

    return (
        f"{format_date_id(start)} "
        f"s.d {format_date_id(end)}"
    )


def count_workdays(start, end):
    start = parse_date(start)
    end = parse_date(end)

    if end < start:
        start, end = end, start

    total = 0
    current = start

    while current <= end:
        if current.weekday() < 5:
            total += 1

        current += timedelta(days=1)

    return total


# ============================================================
# BACKGROUND WORKER
# ============================================================

class LetterWorker(QObject):

    finished = Signal(object)
    error = Signal(str)

    def __init__(self, callback):
        super().__init__()
        self.callback = callback

    def run(self):
        try:
            result = self.callback()
            self.finished.emit(result)

        except Exception:
            error_text = traceback.format_exc()

            print(error_text)

            self.error.emit(error_text)


# ============================================================
# DASHBOARD ICON
# ============================================================

class DashboardIcon(QWidget):

    def __init__(
        self,
        icon_type,
        color="#1769aa",
        parent=None,
    ):
        super().__init__(parent)

        self.icon_type = icon_type
        self.icon_color = QColor(color)

        self.setFixedSize(
            64,
            64
        )

    def paintEvent(self, event):

        painter = QPainter(self)

        painter.setRenderHint(
            QPainter.Antialiasing
        )

        # ----------------------------------------------------
        # LINGKARAN
        # ----------------------------------------------------

        painter.setPen(
            Qt.PenStyle.NoPen
        )

        painter.setBrush(
            QBrush(
                QColor(
                    self.icon_color.red(),
                    self.icon_color.green(),
                    self.icon_color.blue(),
                    30,
                )
            )
        )

        painter.drawEllipse(
            3,
            3,
            58,
            58,
        )

        # ----------------------------------------------------
        # ICON
        # ----------------------------------------------------

        pen = QPen(
            self.icon_color,
            2.5,
        )

        pen.setCapStyle(
            Qt.PenCapStyle.RoundCap
        )

        pen.setJoinStyle(
            Qt.PenJoinStyle.RoundJoin
        )

        painter.setPen(pen)

        painter.setBrush(
            Qt.BrushStyle.NoBrush
        )

        cx = 32

        # ----------------------------------------------------
        # PEGAWAI
        # ----------------------------------------------------

        if self.icon_type == "pegawai":

            painter.drawEllipse(
                24,
                13,
                16,
                16,
            )

            painter.drawArc(
                15,
                31,
                34,
                24,
                0,
                180 * 16,
            )

        # ----------------------------------------------------
        # IZIN
        # ----------------------------------------------------

        elif self.icon_type == "izin":

            painter.drawRoundedRect(
                18,
                10,
                28,
                40,
                3,
                3,
            )

            painter.drawLine(
                24,
                19,
                40,
                19,
            )

            painter.drawLine(
                24,
                27,
                40,
                27,
            )

            painter.drawLine(
                24,
                35,
                35,
                35,
            )

            painter.drawLine(
                32,
                44,
                43,
                33,
            )

        # ----------------------------------------------------
        # SAKIT
        # ----------------------------------------------------

        elif self.icon_type == "sakit":

            painter.drawRoundedRect(
                19,
                10,
                26,
                40,
                3,
                3,
            )

            painter.drawLine(
                cx,
                19,
                cx,
                41,
            )

            painter.drawLine(
                21,
                30,
                43,
                30,
            )

        # ----------------------------------------------------
        # CUTI
        # ----------------------------------------------------

        elif self.icon_type == "cuti":

            painter.drawRoundedRect(
                12,
                18,
                40,
                32,
                5,
                5,
            )

            painter.drawLine(
                12,
                27,
                52,
                27,
            )

            painter.drawLine(
                21,
                12,
                21,
                23,
            )

            painter.drawLine(
                43,
                12,
                43,
                23,
            )

            painter.drawLine(
                23,
                38,
                28,
                43,
            )

            painter.drawLine(
                28,
                43,
                40,
                32,
            )

        # ----------------------------------------------------
        # LEMBUR
        # ----------------------------------------------------

        elif self.icon_type == "lembur":

            painter.drawEllipse(
                14,
                12,
                36,
                36,
            )

            painter.drawLine(
                cx,
                20,
                cx,
                32,
            )

            painter.drawLine(
                cx,
                32,
                40,
                37,
            )

        # ----------------------------------------------------
        # ARSIP
        # ----------------------------------------------------

        elif self.icon_type == "arsip":

            painter.drawRoundedRect(
                11,
                18,
                42,
                32,
                4,
                4,
            )

            painter.drawLine(
                11,
                27,
                53,
                27,
            )

            painter.drawLine(
                20,
                18,
                25,
                12,
            )

            painter.drawLine(
                25,
                12,
                39,
                12,
            )

            painter.drawLine(
                39,
                12,
                44,
                18,
            )

        painter.end()


# ============================================================
# MAIN WINDOW
# ============================================================

class MainWindow(QMainWindow):

    def __init__(self):

        super().__init__()

        self.setWindowTitle(
            "LetterSync - Sistem Surat Menyurat"
        )

        self.resize(
            1250,
            800
        )

        self.thread = None
        self.worker = None
        self.progress = None

        # Lembur worker
        self.worker_thread = None

        self.apply_theme()

        ensure_lembur_table()

        self.setup_ui()

        self.load_employees()

        self.load_archive()

    # ========================================================
    # THEME
    # ========================================================

    def apply_theme(self):

        self.setStyleSheet("""
            /* =================================================
               GLOBAL
               ================================================= */

            QWidget {
                background: #111827;
                color: #E5E7EB;
                font-family: "Segoe UI";
                font-size: 10pt;
            }

            QMainWindow {
                background: #111827;
            }

            /* =================================================
               TAB
               ================================================= */

            QTabWidget::pane {
                background: #151F2E;
                border: 1px solid #2D3A4D;
                border-radius: 10px;
                top: -1px;
            }

            QTabBar::tab {
                background: #1B2636;
                color: #9CAFC3;
                padding: 10px 18px;
                margin-right: 3px;
                border: 1px solid #2D3A4D;
                border-bottom: none;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
            }

            QTabBar::tab:hover {
                background: #243247;
                color: #DCE8F5;
            }

            QTabBar::tab:selected {
                background: #263B52;
                color: #78B7E8;
                border: 1px solid #42617F;
                border-bottom: 2px solid #5B9BD5;
            }

            /* =================================================
               LABEL
               ================================================= */

            QLabel {
                color: #E5E7EB;
                background: transparent;
            }

            /* =================================================
               GROUP BOX
               ================================================= */

            QGroupBox {
                background: #1A2432;
                border: 1px solid #2D3A4D;
                border-radius: 12px;
                margin-top: 16px;
                padding: 18px 12px 12px 12px;
                font-weight: 600;
                color: #DCE8F5;
            }

            QGroupBox::title {
                subcontrol-origin: margin;
                left: 15px;
                padding: 0 9px;
                background: #1A2432;
                color: #76B5E5;
            }

            /* =================================================
               INPUT
               ================================================= */

            QLineEdit,
            QTextEdit,
            QComboBox,
            QDateEdit,
            QTimeEdit,
            QSpinBox {

                background: #202C3D;
                color: #E8EEF5;

                border: 1px solid #3A4A5E;

                border-radius: 8px;

                padding: 8px;

                min-height: 20px;
            }

            QLineEdit:hover,
            QTextEdit:hover,
            QComboBox:hover,
            QDateEdit:hover,
            QTimeEdit:hover,
            QSpinBox:hover {
                border: 1px solid #526A84;
            }

            QLineEdit:focus,
            QTextEdit:focus,
            QComboBox:focus,
            QDateEdit:focus,
            QTimeEdit:focus,
            QSpinBox:focus {

                border: 2px solid #5B9BD5;
                background: #243247;
            }

            QLineEdit:read-only {
                background: #182231;
                color: #9FB0C2;
            }

            /* =================================================
               COMBO BOX
               ================================================= */

            QComboBox QAbstractItemView {
                background: #1B2636;
                color: #E5E7EB;
                border: 1px solid #3A4A5E;
                selection-background-color: #345574;
                selection-color: white;
            }

            /* =================================================
               BUTTON
               ================================================= */

            QPushButton {

                background: #365D80;
                color: #F5F9FC;

                border: 1px solid #47769D;

                border-radius: 8px;

                padding: 9px 16px;

                font-weight: 600;

                min-height: 18px;
            }

            QPushButton:hover {
                background: #42739B;
                border-color: #5B9BD5;
            }

            QPushButton:pressed {
                background: #2B4B68;
            }

            QPushButton:disabled {
                background: #293442;
                color: #68788A;
                border-color: #303C4B;
            }

            /* =================================================
               TABLE
               ================================================= */

            QTableWidget {

                background: #182231;
                color: #E5E7EB;

                border: 1px solid #2D3A4D;

                border-radius: 8px;

                gridline-color: #2B394B;

                selection-background-color: #31516F;

                selection-color: #FFFFFF;

                alternate-background-color: #1D2938;
            }

            QTableWidget::item {
                padding: 7px;
                border-bottom: 1px solid #263445;
            }

            QTableWidget::item:hover {
                background: #26384C;
            }

            QHeaderView::section {

                background: #243247;

                color: #BFD0E1;

                padding: 9px;

                border: none;

                border-bottom: 1px solid #40556D;

                font-weight: 600;
            }

            /* =================================================
               SCROLL BAR
               ================================================= */

            QScrollBar:vertical {
                background: #151F2E;
                width: 10px;
                margin: 2px;
            }

            QScrollBar::handle:vertical {
                background: #3A5068;
                border-radius: 5px;
                min-height: 30px;
            }

            QScrollBar::handle:vertical:hover {
                background: #52728F;
            }

            QScrollBar::add-line:vertical,
            QScrollBar::sub-line:vertical {
                height: 0px;
            }

            QScrollBar:horizontal {
                background: #151F2E;
                height: 10px;
                margin: 2px;
            }

            QScrollBar::handle:horizontal {
                background: #3A5068;
                border-radius: 5px;
                min-width: 30px;
            }

            QScrollBar::handle:horizontal:hover {
                background: #52728F;
            }

            QScrollBar::add-line:horizontal,
            QScrollBar::sub-line:horizontal {
                width: 0px;
            }

            /* =================================================
               CHECK BOX
               ================================================= */

            QCheckBox {
                color: #DCE5EF;
                spacing: 7px;
            }

            /* =================================================
               PROGRESS
               ================================================= */

            QProgressDialog {
                background: #1A2432;
                color: #E5E7EB;
                border: 1px solid #3A4A5E;
            }

            QProgressBar {
                background: #202C3D;
                border: 1px solid #3A4A5E;
                border-radius: 6px;
                text-align: center;
                color: #E5E7EB;
            }

            QProgressBar::chunk {
                background: #5B9BD5;
                border-radius: 5px;
            }

            /* =================================================
               MESSAGE BOX
               ================================================= */

            QMessageBox {
                background: #1A2432;
                color: #E5E7EB;
            }

            QMessageBox QLabel {
                color: #E5E7EB;
            }

            /* =================================================
               STATUS BAR
               ================================================= */

            QStatusBar {
                background: #0B1220;
                color: #AFC2D6;
                padding: 5px 10px;
                border-top: 1px solid #263548;
            }

            /* =================================================
               TOOLTIP
               ================================================= */

            QToolTip {
                background: #26384C;
                color: #F3F6FA;
                border: 1px solid #4B6680;
                padding: 6px;
            }
        """)

    def setup_ui(self):

        central = QWidget()

        self.setCentralWidget(
            central
        )

        main_layout = QVBoxLayout(
            central
        )

        main_layout.setContentsMargins(
            12,
            10,
            12,
            8,
        )

        main_layout.setSpacing(8)

        # ----------------------------------------------------
        # HEADER
        # ----------------------------------------------------

        title = QLabel(
            "LetterSync"
        )

        title.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        title.setStyleSheet("""
            QLabel {
                font-size: 28px;
                font-weight: 800;
                color: #123f68;
                padding: 4px;
            }
        """)

        main_layout.addWidget(
            title
        )

        subtitle = QLabel(
            "Sistem Surat Menyurat • Pembuatan Surat Word dan PDF"
        )

        subtitle.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        subtitle.setStyleSheet("""
            QLabel {
                font-size: 12px;
                color: #718096;
                padding-bottom: 5px;
            }
        """)

        main_layout.addWidget(
            subtitle
        )

        # ----------------------------------------------------
        # TAB
        # ----------------------------------------------------

        self.tabs = QTabWidget()

        self.tabs.setDocumentMode(
            True
        )

        main_layout.addWidget(
            self.tabs
        )

        self.create_dashboard_tab()

        self.create_pegawai_tab()

        self.create_izin_tab()

        self.create_sakit_tab()

        self.create_cuti_tab()

        self.create_lembur_tab()

        self.create_archive_tab()

        # ----------------------------------------------------
        # FOOTER
        # ----------------------------------------------------

        footer = QFrame()

        footer.setStyleSheet("""
            QFrame {
                background: #123f68;
                border-radius: 8px;
            }
        """)

        footer_layout = QHBoxLayout(
            footer
        )

        footer_layout.setContentsMargins(
            14,
            6,
            14,
            6,
        )

        footer_label = QLabel(
            'BK Tech · Support Adha H.Y '
            '<a href="https://saweria.co/simpatik" '
            'style="color:#9bd7ff; font-weight:bold;">'
            'untuk Support & Developer</a>'
        )

        footer_label.setOpenExternalLinks(
            True
        )

        footer_label.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        footer_label.setStyleSheet("""
            QLabel {
                color: #e7f3fb;
                font-size: 9pt;
            }

            QLabel a {
                color: #9bd7ff;
            }
        """)

        footer_layout.addWidget(
            footer_label
        )

        main_layout.addWidget(
            footer
        )

        self.statusBar().showMessage(
            "LetterSync siap digunakan."
        )

    # ========================================================
    # DASHBOARD
    # ========================================================

    def create_dashboard_tab(self):

        tab = QWidget()

        layout = QVBoxLayout(
            tab
        )

        layout.setContentsMargins(
            25,
            20,
            25,
            20,
        )

        layout.setSpacing(
            15
        )

        # ----------------------------------------------------
        # HERO
        # ----------------------------------------------------

        hero = QFrame()

        hero.setMinimumHeight(
            150
        )

        hero.setStyleSheet("""
            QFrame {
                background: #1769aa;
                border-radius: 18px;
            }
        """)

        hero_layout = QHBoxLayout(
            hero
        )

        hero_layout.setContentsMargins(
            25,
            20,
            25,
            20,
        )

        # Icon surat

        hero_icon = QLabel(
            "?"
        )

        hero_icon.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        hero_icon.setFixedSize(
            80,
            80
        )

        hero_icon.setStyleSheet("""
            QLabel {
                background: rgba(255,255,255,0.16);
                color: white;
                border-radius: 40px;
                font-size: 40px;
                font-weight: bold;
            }
        """)

        hero_layout.addWidget(
            hero_icon
        )

        hero_text_layout = QVBoxLayout()

        hero_title = QLabel(
            "LetterSync"
        )

        hero_title.setStyleSheet("""
            QLabel {
                color: white;
                font-size: 27px;
                font-weight: 800;
            }
        """)

        hero_subtitle = QLabel(
            "Sistem Surat Menyurat Terintegrasi"
        )

        hero_subtitle.setStyleSheet("""
            QLabel {
                color: #dcefff;
                font-size: 13px;
            }
        """)

        today_text = QLabel(
            format_date_id(
                date.today()
            )
        )

        today_text.setStyleSheet("""
            QLabel {
                color: #bfe3ff;
                font-size: 11px;
                padding-top: 5px;
            }
        """)

        hero_text_layout.addWidget(
            hero_title
        )

        hero_text_layout.addWidget(
            hero_subtitle
        )

        hero_text_layout.addWidget(
            today_text
        )

        hero_layout.addLayout(
            hero_text_layout
        )

        hero_layout.addStretch()

        layout.addWidget(
            hero
        )

        # ----------------------------------------------------
        # WELCOME
        # ----------------------------------------------------

        welcome = QLabel(
            "Selamat Datang ??"
        )

        welcome.setStyleSheet("""
            QLabel {
                font-size: 21px;
                font-weight: 700;
                color: #243b53;
                padding-top: 2px;
            }
        """)

        layout.addWidget(
            welcome
        )

        info = QLabel(
            "Kelola pegawai, surat izin, surat sakit, surat cuti, "
            "lembur, dan arsip dokumen dalam satu aplikasi."
        )

        info.setStyleSheet("""
            QLabel {
                font-size: 11px;
                color: #718096;
            }
        """)

        layout.addWidget(
            info
        )

        # ----------------------------------------------------
        # STATISTIK
        # ----------------------------------------------------

        stats_layout = QGridLayout()

        stats_layout.setHorizontalSpacing(
            14
        )

        stats_layout.setVerticalSpacing(
            14
        )

        self.dashboard_cards = {}

        cards = [
            (
                "pegawai",
                "pegawai",
                "TOTAL PEGAWAI",
                "#1769aa",
            ),
            (
                "izin",
                "izin",
                "SURAT IZIN",
                "#2e86de",
            ),
            (
                "sakit",
                "sakit",
                "SURAT SAKIT",
                "#e67e22",
            ),
            (
                "cuti",
                "cuti",
                "SURAT CUTI",
                "#27ae60",
            ),
            (
                "lembur",
                "lembur",
                "LEMBUR",
                "#8e44ad",
            ),
            (
                "arsip",
                "arsip",
                "TOTAL ARSIP",
                "#34495e",
            ),
        ]

        for index, (
            key,
            icon_type,
            label,
            color,
        ) in enumerate(cards):

            card = QFrame()

            card.setMinimumHeight(
                145
            )

            card.setStyleSheet("""
                QFrame {
                    background: white;
                    border: 1px solid #dce5ee;
                    border-radius: 14px;
                }
            """)

            shadow = QGraphicsDropShadowEffect()

            shadow.setBlurRadius(
                16
            )

            shadow.setOffset(
                0,
                3
            )

            shadow.setColor(
                QColor(
                    30,
                    70,
                    100,
                    35,
                )
            )

            card.setGraphicsEffect(
                shadow
            )

            card_layout = QHBoxLayout(
                card
            )

            card_layout.setContentsMargins(
                18,
                15,
                18,
                15,
            )

            icon = DashboardIcon(
                icon_type,
                color,
            )

            card_layout.addWidget(
                icon
            )

            text_layout = QVBoxLayout()

            label_widget = QLabel(
                label
            )

            label_widget.setStyleSheet(f"""
                QLabel {{
                    font-size: 10px;
                    font-weight: 700;
                    color: #718096;
                    border: none;
                }}
            """)

            value_widget = QLabel(
                "0"
            )

            value_widget.setStyleSheet(f"""
                QLabel {{
                    font-size: 27px;
                    font-weight: 800;
                    color: {color};
                    border: none;
                }}
            """)

            text_layout.addWidget(
                label_widget
            )

            text_layout.addWidget(
                value_widget
            )

            text_layout.addStretch()

            card_layout.addLayout(
                text_layout
            )

            row = index // 3
            column = index % 3

            stats_layout.addWidget(
                card,
                row,
                column,
            )

            self.dashboard_cards[
                key
            ] = value_widget

        layout.addLayout(
            stats_layout
        )

        # ----------------------------------------------------
        # QUICK ACTION
        # ----------------------------------------------------

        quick_title = QLabel(
            "Akses Cepat"
        )

        quick_title.setStyleSheet("""
            QLabel {
                font-size: 15px;
                font-weight: 700;
                color: #34495e;
                padding-top: 3px;
            }
        """)

        layout.addWidget(
            quick_title
        )

        quick_layout = QHBoxLayout()

        quick_actions = [
            (
                "+  Pegawai",
                1,
            ),
            (
                "??  Surat Izin",
                2,
            ),
            (
                "??  Surat Sakit",
                3,
            ),
            (
                "??  Surat Cuti",
                4,
            ),
            (
                "?  Lembur",
                5,
            ),
            (
                "??  Arsip",
                6,
            ),
        ]

        for text, index in quick_actions:

            button = QPushButton(
                text
            )

            button.setMinimumHeight(
                40
            )

            button.clicked.connect(
                lambda checked=False,
                i=index:
                self.tabs.setCurrentIndex(i)
            )

            quick_layout.addWidget(
                button
            )

        layout.addLayout(
            quick_layout
        )

        # ----------------------------------------------------
        # REFRESH
        # ----------------------------------------------------

        refresh_button = QPushButton(
            "?  Refresh Statistik"
        )

        refresh_button.setMinimumHeight(
            38
        )

        refresh_button.setMaximumWidth(
            220
        )

        refresh_button.clicked.connect(
            self.refresh_dashboard
        )

        refresh_layout = QHBoxLayout()

        refresh_layout.addStretch()

        refresh_layout.addWidget(
            refresh_button
        )

        refresh_layout.addStretch()

        layout.addLayout(
            refresh_layout
        )

        layout.addStretch()

        self.tabs.addTab(
            tab,
            "??  Dashboard"
        )

        self.refresh_dashboard()

    # ========================================================
    # DASHBOARD STATISTICS
    # ========================================================

    def refresh_dashboard(self):

        try:

            conn = get_connection()

            cur = conn.cursor()

            # ------------------------------------------------
            # PEGAWAI
            # ------------------------------------------------

            cur.execute(
                "SELECT COUNT(*) FROM pegawai"
            )

            total_pegawai = cur.fetchone()[0]

            # ------------------------------------------------
            # SURAT
            # ------------------------------------------------

            cur.execute("""
                SELECT jenis, COUNT(*)
                FROM surat
                GROUP BY jenis
            """)

            surat_counts = {}

            for row in cur.fetchall():

                jenis = str(
                    row[0] or ""
                ).strip().lower()

                surat_counts[
                    jenis
                ] = row[1]

            # ------------------------------------------------
            # LEMBUR
            # ------------------------------------------------

            cur.execute(
                "SELECT COUNT(*) FROM lembur"
            )

            total_lembur = cur.fetchone()[0]

            # ------------------------------------------------
            # SURAT
            # ------------------------------------------------

            cur.execute(
                "SELECT COUNT(*) FROM surat"
            )

            total_surat = cur.fetchone()[0]

            conn.close()

            # ------------------------------------------------
            # TAMPILKAN
            # ------------------------------------------------

            self.dashboard_cards[
                "pegawai"
            ].setText(
                str(total_pegawai)
            )

            self.dashboard_cards[
                "izin"
            ].setText(
                str(
                    surat_counts.get(
                        "izin",
                        0,
                    )
                )
            )

            self.dashboard_cards[
                "sakit"
            ].setText(
                str(
                    surat_counts.get(
                        "sakit",
                        0,
                    )
                )
            )

            self.dashboard_cards[
                "cuti"
            ].setText(
                str(
                    surat_counts.get(
                        "cuti",
                        0,
                    )
                )
            )

            self.dashboard_cards[
                "lembur"
            ].setText(
                str(total_lembur)
            )

            self.dashboard_cards[
                "arsip"
            ].setText(
                str(
                    total_surat
                    + total_lembur
                )
            )

        except Exception as e:

            print(
                "Gagal memuat statistik dashboard:",
                e,
            )

    # ========================================================
    # PEGAWAI
    # ========================================================

    def create_pegawai_tab(self):

        tab = QWidget()

        layout = QVBoxLayout(
            tab
        )

        form_box = QGroupBox(
            "Data Pegawai"
        )

        form = QFormLayout(
            form_box
        )

        form.setSpacing(
            10
        )

        self.p_nama = QLineEdit()
        self.p_nip = QLineEdit()
        self.p_jabatan = QLineEdit()
        self.p_unit = QLineEdit()
        self.p_masa = QLineEdit()
        self.p_telepon = QLineEdit()
        self.p_alamat = QTextEdit()

        form.addRow(
            "Nama:",
            self.p_nama,
        )

        form.addRow(
            "NIP:",
            self.p_nip,
        )

        form.addRow(
            "Jabatan:",
            self.p_jabatan,
        )

        form.addRow(
            "Unit Kerja:",
            self.p_unit,
        )

        form.addRow(
            "Masa Kerja:",
            self.p_masa,
        )

        form.addRow(
            "Telepon:",
            self.p_telepon,
        )

        form.addRow(
            "Alamat:",
            self.p_alamat,
        )

        layout.addWidget(
            form_box
        )

        buttons = QHBoxLayout()

        btn_save = QPushButton(
            "+  Simpan Pegawai"
        )

        btn_save.clicked.connect(
            self.save_employee
        )

        btn_refresh = QPushButton(
            "?  Refresh"
        )

        btn_refresh.clicked.connect(
            self.load_employees
        )

        btn_delete = QPushButton(
            "??  Hapus Pegawai"
        )

        btn_delete.clicked.connect(
            self.delete_employee
        )

        buttons.addWidget(
            btn_save
        )

        buttons.addWidget(
            btn_refresh
        )

        buttons.addWidget(
            btn_delete
        )

        buttons.addStretch()

        layout.addLayout(
            buttons
        )

        self.employee_table = QTableWidget()

        self.employee_table.setColumnCount(
            7
        )

        self.employee_table.setHorizontalHeaderLabels([
            "ID",
            "Nama",
            "NIP",
            "Jabatan",
            "Unit Kerja",
            "Masa Kerja",
            "Telepon",
        ])

        self.employee_table.setAlternatingRowColors(
            True
        )

        self.employee_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )

        self.employee_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )

        layout.addWidget(
            self.employee_table
        )

        self.tabs.addTab(
            tab,
            "??  Pegawai"
        )

    # ========================================================
    # LOAD PEGAWAI
    # ========================================================

    def load_employees(self):

        conn = get_connection()

        cur = conn.cursor()

        cur.execute("""
            SELECT
                id,
                nama,
                nip,
                jabatan,
                unit_kerja,
                masa_kerja,
                telepon,
                alamat
            FROM pegawai
            ORDER BY nama
        """)

        self.employees = [
            dict(row)
            for row in cur.fetchall()
        ]

        conn.close()

        if hasattr(
            self,
            "employee_table",
        ):

            self.employee_table.setRowCount(
                len(self.employees)
            )

            for r, emp in enumerate(
                self.employees
            ):

                values = [
                    emp.get(
                        "id",
                        "",
                    ),
                    emp.get(
                        "nama",
                        "",
                    ),
                    emp.get(
                        "nip",
                        "",
                    ),
                    emp.get(
                        "jabatan",
                        "",
                    ),
                    emp.get(
                        "unit_kerja",
                        "",
                    ),
                    emp.get(
                        "masa_kerja",
                        "",
                    ),
                    emp.get(
                        "telepon",
                        "",
                    ),
                ]

                for c, value in enumerate(
                    values
                ):

                    self.employee_table.setItem(
                        r,
                        c,
                        QTableWidgetItem(
                            str(
                                value or ""
                            )
                        ),
                    )

        for combo_name in (
            "i_pegawai",
            "s_pegawai",
            "c_pegawai",
        ):

            if hasattr(
                self,
                combo_name,
            ):

                self.populate_employee_combo(
                    getattr(
                        self,
                        combo_name,
                    )
                )

        # Refresh dashboard
        if hasattr(
            self,
            "dashboard_cards",
        ):
            self.refresh_dashboard()

    # ========================================================
    # POPULATE COMBO
    # ========================================================

    def populate_employee_combo(
        self,
        combo,
    ):

        current_id = combo.currentData()

        combo.blockSignals(
            True
        )

        combo.clear()

        combo.addItem(
            "-- Pilih Pegawai --",
            None,
        )

        for emp in self.employees:

            text = emp.get(
                "nama",
                "",
            )

            if emp.get(
                "nip"
            ):

                text += (
                    f" • "
                    f"{emp['nip']}"
                )

            combo.addItem(
                text,
                emp.get(
                    "id"
                ),
            )

        if current_id is not None:

            index = combo.findData(
                current_id
            )

            if index >= 0:
                combo.setCurrentIndex(
                    index
                )

        combo.blockSignals(
            False
        )

    # ========================================================
    # SAVE PEGAWAI
    # ========================================================

    def save_employee(self):

        nama = self.p_nama.text().strip()

        if not nama:

            QMessageBox.warning(
                self,
                "Data Belum Lengkap",
                "Nama pegawai wajib diisi.",
            )

            return

        conn = get_connection()

        cur = conn.cursor()

        cur.execute("""
            INSERT INTO pegawai
            (
                nama,
                nip,
                jabatan,
                unit_kerja,
                masa_kerja,
                telepon,
                alamat
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            nama,
            self.p_nip.text().strip(),
            self.p_jabatan.text().strip(),
            self.p_unit.text().strip(),
            self.p_masa.text().strip(),
            self.p_telepon.text().strip(),
            self.p_alamat.toPlainText().strip(),
        ))

        conn.commit()

        conn.close()

        self.p_nama.clear()
        self.p_nip.clear()
        self.p_jabatan.clear()
        self.p_unit.clear()
        self.p_masa.clear()
        self.p_telepon.clear()
        self.p_alamat.clear()

        self.load_employees()

        QMessageBox.information(
            self,
            "Berhasil",
            "Data pegawai berhasil disimpan.",
        )

    # ========================================================
    # DELETE PEGAWAI
    # ========================================================

    def delete_employee(self):

        row = self.employee_table.currentRow()

        if row < 0:

            QMessageBox.warning(
                self,
                "Pilih Pegawai",
                "Silakan pilih baris pegawai yang ingin dihapus.",
            )

            return

        id_item = self.employee_table.item(
            row,
            0,
        )

        nama_item = self.employee_table.item(
            row,
            1,
        )

        if not id_item:

            QMessageBox.warning(
                self,
                "Data Tidak Valid",
                "Data pegawai tidak dapat dibaca.",
            )

            return

        employee_id = int(
            id_item.text()
        )

        nama = (
            nama_item.text()
            if nama_item
            else ""
        )

        conn = get_connection()

        try:

            cur = conn.cursor()

            cur.execute(
                """
                SELECT COUNT(*)
                FROM surat
                WHERE pegawai_id = ?
                """,
                (
                    employee_id,
                ),
            )

            jumlah_surat = cur.fetchone()[0]

            cur.execute(
                """
                SELECT COUNT(*)
                FROM lembur
                WHERE pegawai_id = ?
                """,
                (
                    employee_id,
                ),
            )

            jumlah_lembur = cur.fetchone()[0]

        finally:

            conn.close()

        if jumlah_surat > 0 or jumlah_lembur > 0:

            QMessageBox.warning(
                self,
                "Tidak Dapat Dihapus",
                (
                    f"Pegawai '{nama}' sudah digunakan "
                    f"dalam arsip surat/lembur.\n\n"
                    "Data tidak dihapus agar arsip tetap aman."
                ),
            )

            return

        jawaban = QMessageBox.question(
            self,
            "Konfirmasi Hapus Pegawai",
            (
                f"Yakin ingin menghapus pegawai berikut?\n\n"
                f"{nama}\n\n"
                "Data yang sudah dihapus tidak dapat dikembalikan."
            ),
            QMessageBox.StandardButton.Yes
            | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )

        if jawaban != QMessageBox.StandardButton.Yes:
            return

        try:

            conn = get_connection()

            cur = conn.cursor()

            cur.execute(
                """
                DELETE FROM pegawai
                WHERE id = ?
                """,
                (
                    employee_id,
                ),
            )

            conn.commit()

        except Exception as e:

            QMessageBox.critical(
                self,
                "Gagal Menghapus",
                (
                    "Data pegawai tidak dapat dihapus.\n\n"
                    f"{e}"
                ),
            )

            return

        finally:

            try:
                conn.close()
            except Exception:
                pass

        self.load_employees()

        QMessageBox.information(
            self,
            "Berhasil",
            f"Data pegawai '{nama}' berhasil dihapus.",
        )

        self.statusBar().showMessage(
            "Data pegawai berhasil dihapus.",
            5000,
        )

    # ========================================================
    # SURAT IZIN
    # ========================================================

    def create_izin_tab(self):

        tab = QWidget()

        layout = QVBoxLayout(
            tab
        )

        box = QGroupBox(
            "Surat Izin"
        )

        form = QFormLayout(
            box
        )

        self.i_pegawai = QComboBox()

        self.i_tanggal = QLineEdit(
            date.today().strftime(
                "%d-%m-%Y"
            )
        )

        self.i_lama_izin = QLineEdit()

        self.i_lama_izin.setPlaceholderText(
            "Contoh: 1 hari"
        )

        self.i_keperluan = QTextEdit()

        form.addRow(
            "Pegawai:",
            self.i_pegawai,
        )

        form.addRow(
            "Tanggal:",
            self.i_tanggal,
        )

        form.addRow(
            "Lama Izin:",
            self.i_lama_izin,
        )

        form.addRow(
            "Keperluan:",
            self.i_keperluan,
        )

        layout.addWidget(
            box
        )

        self.i_save_button = QPushButton(
            "??  Simpan Surat Izin"
        )

        self.i_save_button.clicked.connect(
            self.save_izin
        )

        layout.addWidget(
            self.i_save_button
        )

        layout.addStretch()

        self.tabs.addTab(
            tab,
            "??  Surat Izin"
        )

    def save_izin(self):

        employee = self.get_selected_employee(
            self.i_pegawai
        )

        if not employee:

            QMessageBox.warning(
                self,
                "Data Belum Lengkap",
                "Silakan pilih pegawai.",
            )

            return

        keperluan = (
            self.i_keperluan
            .toPlainText()
            .strip()
        )

        if not keperluan:

            QMessageBox.warning(
                self,
                "Data Belum Lengkap",
                "Keperluan wajib diisi.",
            )

            return

        lama_izin = (
            self.i_lama_izin
            .text()
            .strip()
        )

        if not lama_izin:

            QMessageBox.warning(
                self,
                "Data Belum Lengkap",
                "Lama izin wajib diisi.",
            )

            self.i_lama_izin.setFocus()

            return

        try:

            tanggal = parse_date(
                self.i_tanggal.text()
            )

        except Exception:

            QMessageBox.warning(
                self,
                "Tanggal Salah",
                "Gunakan format tanggal DD-MM-YYYY.",
            )

            return

        nama = employee.get(
            "nama",
            "",
        )

        nip = employee.get(
            "nip",
            "",
        )

        jabatan = employee.get(
            "jabatan",
            "",
        )

        hari = format_day_id(
            tanggal
        )

        tanggal_izin = format_date_id(
            tanggal
        )

        values = {

            "{{NAMA}}":
                nama,

            "{{NIP}}":
                nip,

            "{{JABATAN}}":
                jabatan,

            "{{HARI}}":
                hari,

            "{{TANGGAL_IZIN}}":
                tanggal_izin,

            "{{LAMA_IZIN}}":
                lama_izin,

            "{{KEPERLUAN}}":
                keperluan,

            "{{TANGGAL_SURAT}}":
                tanggal_izin,
        }

        self.start_letter_worker(
            "Surat Izin",
            "surat_izin.docx",
            employee,
            tanggal,
            values,
        )


    # ========================================================
    # SURAT SAKIT
    # ========================================================

    def create_sakit_tab(self):

        tab = QWidget()

        layout = QVBoxLayout(
            tab
        )

        box = QGroupBox(
            "Surat Sakit"
        )

        form = QFormLayout(
            box
        )

        self.s_pegawai = QComboBox()

        self.s_mulai = QLineEdit(
            date.today().strftime(
                "%d-%m-%Y"
            )
        )

        self.s_selesai = QLineEdit(
            date.today().strftime(
                "%d-%m-%Y"
            )
        )

        self.s_keperluan = QTextEdit()

        form.addRow(
            "Pegawai:",
            self.s_pegawai,
        )

        form.addRow(
            "Tanggal Mulai:",
            self.s_mulai,
        )

        form.addRow(
            "Tanggal Selesai:",
            self.s_selesai,
        )

        form.addRow(
            "Keperluan:",
            self.s_keperluan,
        )

        layout.addWidget(
            box
        )

        self.s_save_button = QPushButton(
            "??  Simpan Surat Sakit"
        )

        self.s_save_button.clicked.connect(
            self.save_sakit
        )

        layout.addWidget(
            self.s_save_button
        )

        layout.addStretch()

        self.tabs.addTab(
            tab,
            "??  Surat Sakit"
        )

    def save_sakit(self):

        employee = self.get_selected_employee(
            self.s_pegawai
        )

        if not employee:

            QMessageBox.warning(
                self,
                "Data Belum Lengkap",
                "Silakan pilih pegawai.",
            )

            return

        keperluan = (
            self.s_keperluan
            .toPlainText()
            .strip()
        )

        if not keperluan:

            QMessageBox.warning(
                self,
                "Data Belum Lengkap",
                "Keperluan wajib diisi.",
            )

            return

        try:

            mulai = parse_date(
                self.s_mulai.text()
            )

            selesai = parse_date(
                self.s_selesai.text()
            )

        except Exception:

            QMessageBox.warning(
                self,
                "Tanggal Salah",
                "Gunakan format DD-MM-YYYY.",
            )

            return

        if selesai < mulai:

            QMessageBox.warning(
                self,
                "Tanggal Salah",
                "Tanggal selesai tidak boleh sebelum tanggal mulai.",
            )

            return

        jumlah_hari = count_workdays(
            mulai,
            selesai,
        )

        if jumlah_hari <= 0:

            QMessageBox.warning(
                self,
                "Tanggal Tidak Valid",
                "Rentang tanggal tersebut tidak memiliki hari kerja.",
            )

            return

        values = {

            "Hariyadi Adha, S.Pd":
                employee.get(
                    "nama",
                    "",
                ),

            "07012102016113030":
                employee.get(
                    "nip",
                    "",
                ),

            "Petugas Loket BPJS":
                employee.get(
                    "jabatan",
                    "",
                ),

            "Kamis s.d Jumat":
                format_range_day(
                    mulai,
                    selesai,
                ),

            "16 Oktober 2025 s.d 17 Oktober 2025":
                format_range_date(
                    mulai,
                    selesai,
                ),

            "2 hari kerja":
                f"{jumlah_hari} hari kerja",

            "Menjaga Istri sakit":
                keperluan,
        }

        self.start_letter_worker(
            "Surat Sakit",
            "surat_sakit.docx",
            employee,
            mulai,
            values,
        )

    # ========================================================
    # SURAT CUTI
    # ========================================================

    def create_cuti_tab(self):

        tab = QWidget()

        layout = QVBoxLayout(
            tab
        )

        box = QGroupBox(
            "Surat Cuti"
        )

        form = QFormLayout(
            box
        )

        self.c_pegawai = QComboBox()

        self.c_lama = QLineEdit()

        self.c_alamat = QLineEdit()

        self.c_alasan = QTextEdit()

        self.c_masa = QLineEdit()

        self.c_telepon = QLineEdit()

        self.c_sisa_n1 = QLineEdit()

        self.c_sisa_n2 = QLineEdit()

        form.addRow(
            "Pegawai:",
            self.c_pegawai,
        )

        form.addRow(
            "Lama Cuti:",
            self.c_lama,
        )

        form.addRow(
            "Alamat Cuti:",
            self.c_alamat,
        )

        form.addRow(
            "Alasan Cuti:",
            self.c_alasan,
        )

        form.addRow(
            "Masa Kerja:",
            self.c_masa,
        )

        form.addRow(
            "Telepon:",
            self.c_telepon,
        )

        form.addRow(
            "Sisa N1:",
            self.c_sisa_n1,
        )

        form.addRow(
            "Sisa N2:",
            self.c_sisa_n2,
        )

        layout.addWidget(
            box
        )

        self.c_save_button = QPushButton(
            "??  Simpan Surat Cuti"
        )

        self.c_save_button.clicked.connect(
            self.save_cuti
        )

        layout.addWidget(
            self.c_save_button
        )

        layout.addStretch()

        self.c_pegawai.currentIndexChanged.connect(
            self.on_cuti_employee_changed
        )

        self.tabs.addTab(
            tab,
            "??  Surat Cuti"
        )

    def on_cuti_employee_changed(self):

        employee = self.get_selected_employee(
            self.c_pegawai
        )

        if not employee:
            return

        self.c_masa.setText(
            employee.get(
                "masa_kerja"
            ) or ""
        )

        self.c_telepon.setText(
            employee.get(
                "telepon"
            ) or ""
        )

        self.c_alamat.setText(
            employee.get(
                "alamat"
            ) or ""
        )

    def save_cuti(self):

        employee = self.get_selected_employee(
            self.c_pegawai
        )

        if not employee:

            QMessageBox.warning(
                self,
                "Data Belum Lengkap",
                "Silakan pilih pegawai.",
            )

            return

        values = {

            "{{NAMA}}":
                employee.get(
                    "nama",
                    "",
                ),

            "{{NIP}}":
                employee.get(
                    "nip",
                    "",
                ),

            "{{ALAMAT_CUTI}}":
                self.c_alamat
                .text()
                .strip(),

            "{{ALASAN_CUTI}}":
                self.c_alasan
                .toPlainText()
                .strip(),

            "{{LAMA_CUTI}}":
                self.c_lama
                .text()
                .strip(),

            "{{MASA_KERJA}}":
                self.c_masa
                .text()
                .strip(),

            "{{TELEPON}}":
                self.c_telepon
                .text()
                .strip(),

            "{{SISA_N1}}":
                self.c_sisa_n1
                .text()
                .strip(),

            "{{SISA_N2}}":
                self.c_sisa_n2
                .text()
                .strip(),
        }

        self.start_letter_worker(
            "Surat Cuti",
            "surat_cuti.docx",
            employee,
            date.today(),
            values,
        )

    # ========================================================
    # GET EMPLOYEE
    # ========================================================

    def get_selected_employee(
        self,
        combo,
    ):

        employee_id = combo.currentData()

        if employee_id is None:
            return None

        for employee in self.employees:

            if employee.get(
                "id"
            ) == employee_id:

                return dict(
                    employee
                )

        return None

    # ========================================================
    # BACKGROUND SAVE
    # ========================================================

    def start_letter_worker(
        self,
        jenis,
        template_name,
        employee,
        tanggal,
        values,
    ):

        timestamp = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )

        safe_name = self.safe_filename(
            employee.get(
                "nama"
            ) or "pegawai"
        )

        output_name = (
            f"{jenis.replace(' ', '_')}_"
            f"{safe_name}_"
            f"{timestamp}.docx"
        )

        self.set_save_buttons_enabled(
            False
        )

        self.progress = QProgressDialog(
            "Sedang membuat surat...",
            None,
            0,
            0,
            self,
        )

        self.progress.setWindowTitle(
            "LetterSync"
        )

        self.progress.setCancelButton(
            None
        )

        self.progress.setMinimumDuration(
            0
        )

        self.progress.setWindowModality(
            Qt.WindowModality.WindowModal
        )

        self.progress.show()

        self.statusBar().showMessage(
            "Sedang membuat surat..."
        )

        QApplication.processEvents()

        def do_work():

            render_template(
                template_name,
                output_name,
                values,
            )

            docx_path = (
                WORD_DIR
                / output_name
            )

            if not docx_path.exists():

                raise FileNotFoundError(
                    "File Word tidak ditemukan:\n"
                    f"{docx_path}"
                )

            create_pdf_from_docx(
                docx_path
            )

            pdf_path = (
                PDF_DIR
                / f"{docx_path.stem}.pdf"
            )

            if not pdf_path.exists():

                raise FileNotFoundError(
                    "PDF gagal dibuat.\n\n"
                    "Pastikan LibreOffice tersedia."
                )

            conn = get_connection()

            cur = conn.cursor()

            cur.execute("""
                INSERT INTO surat
                (
                    jenis,
                    nomor,
                    pegawai_id,
                    tanggal_surat,
                    status,
                    file_path,
                    pdf_path,
                    created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                jenis,
                "",
                employee.get(
                    "id"
                ),
                tanggal.isoformat(),
                "Selesai",
                str(
                    docx_path
                ),
                str(
                    pdf_path
                ),
                datetime.now().isoformat(
                    timespec="seconds"
                ),
            ))

            conn.commit()

            surat_id = cur.lastrowid

            conn.close()

            return {
                "id": surat_id,
                "jenis": jenis,
                "employee":
                    employee.get(
                        "nama",
                        "",
                    ),
                "word":
                    str(
                        docx_path
                    ),
                "pdf":
                    str(
                        pdf_path
                    ),
            }

        self.thread = QThread()

        self.worker = LetterWorker(
            do_work
        )

        self.worker.moveToThread(
            self.thread
        )

        self.thread.started.connect(
            self.worker.run
        )

        self.worker.finished.connect(
            self.letter_finished
        )

        self.worker.error.connect(
            self.letter_error
        )

        self.worker.finished.connect(
            self.thread.quit
        )

        self.worker.error.connect(
            self.thread.quit
        )

        self.thread.finished.connect(
            self.worker.deleteLater
        )

        self.thread.finished.connect(
            self.thread.deleteLater
        )

        self.thread.start()

    # ========================================================
    # LETTER SUCCESS
    # ========================================================

    def letter_finished(
        self,
        result,
    ):

        if self.progress:

            self.progress.close()

            self.progress.deleteLater()

            self.progress = None

        self.set_save_buttons_enabled(
            True
        )

        self.statusBar().showMessage(
            "Surat berhasil dibuat.",
            7000,
        )

        self.load_archive()

        word_path = result.get(
            "word",
            "",
        )

        pdf_path = result.get(
            "pdf",
            "",
        )

        message = (
            "Surat berhasil dibuat.\n\n"
            f"Pegawai : {result.get('employee', '')}\n"
            f"Jenis   : {result.get('jenis', '')}\n\n"
            f"Word:\n{word_path}\n\n"
            f"PDF:\n{pdf_path}"
        )

        box = QMessageBox(
            self
        )

        box.setWindowTitle(
            "Surat Berhasil Dibuat"
        )

        box.setText(
            message
        )

        open_word = box.addButton(
            "Buka Word",
            QMessageBox.ButtonRole.AcceptRole,
        )

        open_pdf = box.addButton(
            "Buka PDF",
            QMessageBox.ButtonRole.AcceptRole,
        )

        open_folder = box.addButton(
            "Buka Folder",
            QMessageBox.ButtonRole.AcceptRole,
        )

        box.addButton(
            "Tutup",
            QMessageBox.ButtonRole.RejectRole,
        )

        box.exec()

        clicked = box.clickedButton()

        if clicked == open_word:

            self.open_file(
                word_path
            )

        elif clicked == open_pdf:

            self.open_file(
                pdf_path
            )

        elif clicked == open_folder:

            self.open_folder(
                str(
                    Path(
                        pdf_path
                    ).parent
                )
            )

    # ========================================================
    # LETTER ERROR
    # ========================================================

    def letter_error(
        self,
        message,
    ):

        if self.progress:

            self.progress.close()

            self.progress.deleteLater()

            self.progress = None

        self.set_save_buttons_enabled(
            True
        )

        self.statusBar().showMessage(
            "Gagal membuat surat.",
            7000,
        )

        QMessageBox.critical(
            self,
            "Gagal Membuat Surat",
            "Surat tidak berhasil dibuat.\n\n"
            + message,
        )

    # ========================================================
    # ENABLE / DISABLE BUTTON
    # ========================================================

    def set_save_buttons_enabled(
        self,
        enabled,
    ):

        for name in (
            "i_save_button",
            "s_save_button",
            "c_save_button",
            "l_save_button",
        ):

            if hasattr(
                self,
                name,
            ):

                getattr(
                    self,
                    name,
                ).setEnabled(
                    enabled
                )

    # ========================================================
    # SAFE FILENAME
    # ========================================================

    def safe_filename(
        self,
        text,
    ):

        invalid = '<>:"/\\|?*'

        result = str(
            text
        )

        for char in invalid:

            result = result.replace(
                char,
                "_",
            )

        return (
            result.strip()
            or "pegawai"
        )

    # ========================================================
    # LEMBUR
    # ========================================================

    def create_lembur_tab(self):

        self.lembur_tab = QWidget()

        layout = QVBoxLayout(
            self.lembur_tab
        )

        layout.setContentsMargins(
            18,
            18,
            18,
            18,
        )

        layout.setSpacing(
            14
        )

        # ----------------------------------------------------
        # TITLE
        # ----------------------------------------------------

        title = QLabel(
            "?  LEMBUR"
        )

        title.setStyleSheet("""
            QLabel {
                font-size: 20pt;
                font-weight: bold;
                color: #123f68;
                padding-bottom: 4px;
            }
        """)

        layout.addWidget(
            title
        )

        subtitle = QLabel(
            "Buat satu surat lembur untuk beberapa tanggal. "
            "Jam dan total lembur dihitung otomatis untuk setiap tanggal."
        )

        subtitle.setStyleSheet("""
            QLabel {
                color: #60757a;
                font-size: 10pt;
                padding-bottom: 8px;
            }
        """)

        layout.addWidget(
            subtitle
        )

        # ----------------------------------------------------
        # DATA PEGAWAI
        # ----------------------------------------------------

        form_box = QGroupBox(
            "Data Lembur"
        )

        form = QFormLayout(
            form_box
        )

        form.setContentsMargins(
            18,
            18,
            18,
            18,
        )

        form.setHorizontalSpacing(
            18
        )

        form.setVerticalSpacing(
            12
        )

        self.lembur_employee_combo = QComboBox()

        self.lembur_employee_combo.setMinimumHeight(
            38
        )

        self.lembur_employee_combo.setMinimumWidth(
            350
        )

        conn = get_connection()

        rows = conn.execute("""
            SELECT
                id,
                nama,
                nip,
                jabatan
            FROM pegawai
            ORDER BY nama
        """).fetchall()

        conn.close()

        self.lembur_employee_combo.addItem(
            "-- Pilih Pegawai --",
            None,
        )

        for row in rows:

            nama = row["nama"] or ""

            nip = row["nip"] or ""

            display = (
                f"{nama} • {nip}"
                if nip
                else nama
            )

            self.lembur_employee_combo.addItem(
                display,
                row["id"],
            )

        self.lembur_employee_combo.currentIndexChanged.connect(
            self.update_lembur_employee_info
        )

        form.addRow(
            "Pegawai",
            self.lembur_employee_combo,
        )

        self.lembur_employee_name = QLineEdit()

        self.lembur_employee_name.setReadOnly(
            True
        )

        self.lembur_employee_name.setPlaceholderText(
            "Otomatis mengikuti pegawai yang dipilih"
        )

        form.addRow(
            "Nama Pelaksana",
            self.lembur_employee_name,
        )

        self.lembur_employee_nip = QLineEdit()

        self.lembur_employee_nip.setReadOnly(
            True
        )

        self.lembur_employee_nip.setPlaceholderText(
            "Otomatis mengikuti pegawai yang dipilih"
        )

        form.addRow(
            "NIP / NIPK Pelaksana",
            self.lembur_employee_nip,
        )

        self.lembur_letter_date = QDateEdit()

        self.lembur_letter_date.setCalendarPopup(
            True
        )

        self.lembur_letter_date.setDate(
            QDate.currentDate()
        )

        self.lembur_letter_date.setDisplayFormat(
            "dd/MM/yyyy"
        )

        self.lembur_letter_date.setMinimumHeight(
            38
        )

        self.lembur_letter_date.setMinimumWidth(
            180
        )

        self.lembur_letter_date.setKeyboardTracking(
            False
        )

        form.addRow(
            "Tanggal Pembuatan Surat",
            self.lembur_letter_date,
        )

        layout.addWidget(
            form_box
        )

        # ----------------------------------------------------
        # INPUT TANGGAL
        # ----------------------------------------------------

        input_box = QGroupBox(
            "Tambah Tanggal Lembur"
        )

        input_layout = QVBoxLayout(
            input_box
        )

        input_layout.setContentsMargins(
            18,
            18,
            18,
            18,
        )

        input_layout.setSpacing(
            10
        )

        row_input = QHBoxLayout()

        row_input.setSpacing(
            10
        )

        date_label = QLabel(
            "Tanggal"
        )

        self.lembur_date = QDateEdit()

        self.lembur_date.setCalendarPopup(
            True
        )

        self.lembur_date.setDate(
            QDate.currentDate()
        )

        self.lembur_date.setDisplayFormat(
            "dd/MM/yyyy"
        )

        self.lembur_date.setMinimumHeight(
            38
        )

        self.lembur_date.setMinimumWidth(
            145
        )

        start_label = QLabel(
            "Mulai"
        )

        self.lembur_start = QTimeEdit()

        self.lembur_start.setDisplayFormat(
            "HH:mm"
        )

        self.lembur_start.setTime(
            QTime(
                17,
                0,
            )
        )

        self.lembur_start.setMinimumHeight(
            38
        )

        self.lembur_start.setMinimumWidth(
            100
        )

        end_label = QLabel(
            "Selesai"
        )

        self.lembur_end = QTimeEdit()

        self.lembur_end.setDisplayFormat(
            "HH:mm"
        )

        self.lembur_end.setTime(
            QTime(
                21,
                0,
            )
        )

        self.lembur_end.setMinimumHeight(
            38
        )

        self.lembur_end.setMinimumWidth(
            100
        )

        total_label = QLabel(
            "Durasi"
        )

        self.lembur_total = QLineEdit()

        self.lembur_total.setReadOnly(
            True
        )

        self.lembur_total.setMinimumHeight(
            38
        )

        self.lembur_total.setMinimumWidth(
            90
        )

        self.lembur_total.setPlaceholderText(
            "Otomatis"
        )

        row_input.addWidget(
            date_label
        )

        row_input.addWidget(
            self.lembur_date
        )

        row_input.addWidget(
            start_label
        )

        row_input.addWidget(
            self.lembur_start
        )

        row_input.addWidget(
            end_label
        )

        row_input.addWidget(
            self.lembur_end
        )

        row_input.addWidget(
            total_label
        )

        row_input.addWidget(
            self.lembur_total
        )

        input_layout.addLayout(
            row_input
        )

        report_label = QLabel(
            "Laporan Lembur"
        )

        self.lembur_report = QTextEdit()

        self.lembur_report.setMinimumHeight(
            80
        )

        self.lembur_report.setPlaceholderText(
            "Masukkan laporan/kegiatan lembur untuk tanggal ini."
        )

        input_layout.addWidget(
            report_label
        )

        input_layout.addWidget(
            self.lembur_report
        )

        add_layout = QHBoxLayout()

        add_layout.addStretch()

        self.lembur_add_button = QPushButton(
            "+  Tambah Tanggal Lembur"
        )

        self.lembur_add_button.setMinimumHeight(
            40
        )

        self.lembur_add_button.setMinimumWidth(
            220
        )

        self.lembur_add_button.clicked.connect(
            self.add_lembur_item
        )

        add_layout.addWidget(
            self.lembur_add_button
        )

        input_layout.addLayout(
            add_layout
        )

        layout.addWidget(
            input_box
        )

        # ----------------------------------------------------
        # TABLE
        # ----------------------------------------------------

        list_box = QGroupBox(
            "Daftar Tanggal Lembur"
        )

        list_layout = QVBoxLayout(
            list_box
        )

        list_layout.setContentsMargins(
            18,
            18,
            18,
            18,
        )

        self.lembur_items_table = QTableWidget()

        self.lembur_items_table.setColumnCount(
            7
        )

        self.lembur_items_table.setHorizontalHeaderLabels([
            "No",
            "Tanggal",
            "Jam Mulai",
            "Jam Selesai",
            "Total",
            "Laporan",
            "Aksi",
        ])

        self.lembur_items_table.setMinimumHeight(
            220
        )

        self.lembur_items_table.setAlternatingRowColors(
            True
        )

        self.lembur_items_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )

        self.lembur_items_table.setEditTriggers(
            QTableWidget.EditTrigger.NoEditTriggers
        )

        header = (
            self.lembur_items_table
            .horizontalHeader()
        )

        header.setSectionResizeMode(
            0,
            QHeaderView.ResizeMode.ResizeToContents,
        )

        header.setSectionResizeMode(
            1,
            QHeaderView.ResizeMode.ResizeToContents,
        )

        header.setSectionResizeMode(
            2,
            QHeaderView.ResizeMode.ResizeToContents,
        )

        header.setSectionResizeMode(
            3,
            QHeaderView.ResizeMode.ResizeToContents,
        )

        header.setSectionResizeMode(
            4,
            QHeaderView.ResizeMode.ResizeToContents,
        )

        header.setSectionResizeMode(
            5,
            QHeaderView.ResizeMode.Stretch,
        )

        header.setSectionResizeMode(
            6,
            QHeaderView.ResizeMode.ResizeToContents,
        )

        list_layout.addWidget(
            self.lembur_items_table
        )

        total_layout = QHBoxLayout()

        total_layout.addStretch()

        total_text = QLabel(
            "TOTAL SELURUH LEMBUR:"
        )

        total_text.setStyleSheet("""
            QLabel {
                font-weight: bold;
                font-size: 10pt;
                color: #123f68;
            }
        """)

        self.lembur_grand_total = QLabel(
            "0 jam"
        )

        self.lembur_grand_total.setStyleSheet("""
            QLabel {
                font-weight: bold;
                font-size: 13pt;
                color: #1769aa;
            }
        """)

        total_layout.addWidget(
            total_text
        )

        total_layout.addWidget(
            self.lembur_grand_total
        )

        list_layout.addLayout(
            total_layout
        )

        layout.addWidget(
            list_box
        )

        # ----------------------------------------------------
        # SAVE
        # ----------------------------------------------------

        button_layout = QHBoxLayout()

        button_layout.addStretch()

        self.l_save_button = QPushButton(
            "?  Buat Surat Lembur"
        )

        self.l_save_button.setMinimumHeight(
            42
        )

        self.l_save_button.setMinimumWidth(
            210
        )

        self.l_save_button.clicked.connect(
            self.save_lembur
        )

        button_layout.addWidget(
            self.l_save_button
        )

        layout.addLayout(
            button_layout
        )

        layout.addStretch()

        # ----------------------------------------------------
        # INTERNAL DATA
        # ----------------------------------------------------

        self.lembur_items = []

        self.lembur_start.timeChanged.connect(
            self.update_lembur_duration
        )

        self.lembur_end.timeChanged.connect(
            self.update_lembur_duration
        )

        self.update_lembur_duration()

        self.tabs.addTab(
            self.lembur_tab,
            "?  Lembur"
        )

    # ========================================================
    # ADD LEMBUR ITEM
    # ========================================================

    def add_lembur_item(self):

        employee_id = (
            self.lembur_employee_combo
            .currentData()
        )

        if not employee_id:

            QMessageBox.warning(
                self,
                "Data Belum Lengkap",
                "Silakan pilih pegawai terlebih dahulu.",
            )

            return

        qdate = (
            self.lembur_date.date()
        )

        tanggal_iso = qdate.toString(
            "yyyy-MM-dd"
        )

        tanggal_display = qdate.toString(
            "dd/MM/yyyy"
        )

        start = (
            self.lembur_start.time()
        )

        end = (
            self.lembur_end.time()
        )

        start_minutes = (
            start.hour() * 60
            + start.minute()
        )

        end_minutes = (
            end.hour() * 60
            + end.minute()
        )

        if end_minutes >= start_minutes:

            duration_minutes = (
                end_minutes
                - start_minutes
            )

        else:

            duration_minutes = (
                24 * 60
                - start_minutes
                + end_minutes
            )

        if duration_minutes <= 0:

            QMessageBox.warning(
                self,
                "Jam Tidak Valid",
                "Jam selesai lembur harus berbeda dari jam mulai.",
            )

            return

        for item in self.lembur_items:

            if item.get(
                "tanggal_iso"
            ) == tanggal_iso:

                QMessageBox.warning(
                    self,
                    "Tanggal Sudah Ada",
                    (
                        f"Tanggal {tanggal_display} "
                        "sudah ada dalam daftar lembur."
                    ),
                )

                return

        jam = (
            duration_minutes
            // 60
        )

        menit = (
            duration_minutes
            % 60
        )

        if menit:

            total_text = (
                f"{jam} jam "
                f"{menit} menit"
            )

        else:

            total_text = (
                f"{jam} jam"
            )

        jam_mulai = start.toString(
            "HH:mm"
        )

        jam_selesai = end.toString(
            "HH:mm"
        )

        rentang = (
            f"{jam_mulai} s.d "
            f"{jam_selesai}"
        )

        laporan = (
            self.lembur_report
            .toPlainText()
            .strip()
        )

        item = {

            "tanggal":
                tanggal_display,

            "tanggal_iso":
                tanggal_iso,

            "jam_mulai":
                jam_mulai,

            "jam_selesai":
                jam_selesai,

            "rentang_jam":
                rentang,

            "total":
                total_text,

            "total_menit":
                duration_minutes,

            "laporan":
                laporan,
        }

        self.lembur_items.append(
            item
        )

        self.refresh_lembur_items_table()

        self.lembur_date.setDate(
            qdate.addDays(1)
        )

        self.lembur_report.clear()

    # ========================================================
    # REMOVE LEMBUR
    # ========================================================

    def remove_lembur_item(
        self,
        index,
    ):

        if (
            index < 0
            or index >= len(
                self.lembur_items
            )
        ):
            return

        item = self.lembur_items[
            index
        ]

        tanggal = item.get(
            "tanggal",
            "",
        )

        reply = QMessageBox.question(
            self,
            "Hapus Tanggal",
            f"Hapus tanggal lembur {tanggal}?",
            QMessageBox.StandardButton.Yes
            | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )

        if reply != QMessageBox.StandardButton.Yes:
            return

        self.lembur_items.pop(
            index
        )

        self.refresh_lembur_items_table()

    # ========================================================
    # REFRESH LEMBUR TABLE
    # ========================================================

    def refresh_lembur_items_table(
        self
    ):

        self.lembur_items_table.setRowCount(
            0
        )

        total_minutes = 0

        for index, item in enumerate(
            self.lembur_items
        ):

            row = (
                self.lembur_items_table
                .rowCount()
            )

            self.lembur_items_table.insertRow(
                row
            )

            values = [
                str(index + 1),
                item.get(
                    "tanggal",
                    "",
                ),
                item.get(
                    "jam_mulai",
                    "",
                ),
                item.get(
                    "jam_selesai",
                    "",
                ),
                item.get(
                    "total",
                    "",
                ),
            ]

            for column, value in enumerate(
                values
            ):

                self.lembur_items_table.setItem(
                    row,
                    column,
                    QTableWidgetItem(
                        str(value)
                    ),
                )

            laporan = item.get(
                "laporan",
                "",
            )

            if not laporan:
                laporan = "-"

            self.lembur_items_table.setItem(
                row,
                5,
                QTableWidgetItem(
                    laporan
                ),
            )

            delete_button = QPushButton(
                "Hapus"
            )

            delete_button.setMinimumHeight(
                30
            )

            delete_button.clicked.connect(
                lambda checked=False,
                i=index:
                self.remove_lembur_item(i)
            )

            self.lembur_items_table.setCellWidget(
                row,
                6,
                delete_button
            )

            total_minutes += int(
                item.get(
                    "total_menit",
                    0,
                )
                or 0
            )

        total_hours = (
            total_minutes
            // 60
        )

        total_remainder = (
            total_minutes
            % 60
        )

        if total_remainder:

            total_text = (
                f"{total_hours} jam "
                f"{total_remainder} menit"
            )

        else:

            total_text = (
                f"{total_hours} jam"
            )

        self.lembur_grand_total.setText(
            total_text
        )

    # ========================================================
    # UPDATE DURASI
    # ========================================================

    def update_lembur_duration(
        self
    ):

        start = (
            self.lembur_start.time()
        )

        end = (
            self.lembur_end.time()
        )

        start_minutes = (
            start.hour() * 60
            + start.minute()
        )

        end_minutes = (
            end.hour() * 60
            + end.minute()
        )

        if end_minutes >= start_minutes:

            duration_minutes = (
                end_minutes
                - start_minutes
            )

        else:

            duration_minutes = (
                24 * 60
                - start_minutes
                + end_minutes
            )

        if duration_minutes <= 0:

            self.lembur_total.setText(
                ""
            )

            return

        jam = (
            duration_minutes
            // 60
        )

        menit = (
            duration_minutes
            % 60
        )

        if menit:

            text = (
                f"{jam} jam "
                f"{menit} menit"
            )

        else:

            text = (
                f"{jam} jam"
            )

        self.lembur_total.setText(
            text
        )

    # ========================================================
    # UPDATE PEGAWAI LEMBUR
    # ========================================================

    def update_lembur_employee_info(
        self
    ):

        employee_id = (
            self.lembur_employee_combo
            .currentData()
        )

        self.lembur_employee_name.clear()

        self.lembur_employee_nip.clear()

        if not employee_id:
            return

        conn = get_connection()

        cur = conn.cursor()

        cur.execute(
            """
            SELECT
                nama,
                nip
            FROM pegawai
            WHERE id = ?
            """,
            (
                employee_id,
            ),
        )

        row = cur.fetchone()

        conn.close()

        if row:

            self.lembur_employee_name.setText(
                row[0] or ""
            )

            self.lembur_employee_nip.setText(
                row[1] or ""
            )

    # ========================================================
    # SAVE LEMBUR
    # ========================================================

    def save_lembur(self):

        employee_id = (
            self.lembur_employee_combo
            .currentData()
        )

        if not employee_id:

            QMessageBox.warning(
                self,
                "Data Belum Lengkap",
                "Silakan pilih pegawai terlebih dahulu.",
            )

            return

        if not self.lembur_items:

            QMessageBox.warning(
                self,
                "Tanggal Lembur Belum Ada",
                (
                    "Tambahkan minimal satu tanggal "
                    "lembur terlebih dahulu."
                ),
            )

            return

        tanggal_surat = (
            self.lembur_letter_date.date()
        )

        tanggal_surat_iso = (
            tanggal_surat.toString(
                "yyyy-MM-dd"
            )
        )

        tanggal_surat_text = (
            f"{tanggal_surat.day()} "
            f"{BULAN_ID[tanggal_surat.month() - 1]} "
            f"{tanggal_surat.year()}"
        )

        conn = get_connection()

        row = conn.execute(
            """
            SELECT *
            FROM pegawai
            WHERE id = ?
            """,
            (
                employee_id,
            ),
        ).fetchone()

        conn.close()

        if not row:

            QMessageBox.warning(
                self,
                "Data Tidak Ditemukan",
                "Data pegawai tidak ditemukan.",
            )

            return

        nama = row["nama"] or ""

        nip = row["nip"] or ""

        jabatan = row["jabatan"] or ""

        unit_kerja = row["unit_kerja"] or ""

        template_items = []

        for item in self.lembur_items:

            iso = item.get(
                "tanggal_iso",
                "",
            )

            qdate = QDate.fromString(
                iso,
                "yyyy-MM-dd"
            )

            if qdate.isValid():

                tanggal_text = (
                    f"{qdate.day()} "
                    f"{BULAN_ID[qdate.month() - 1]} "
                    f"{qdate.year()}"
                )

            else:

                tanggal_text = item.get(
                    "tanggal",
                    "",
                )

            template_items.append({

                "tanggal":
                    tanggal_text,

                "tanggal_iso":
                    iso,

                "jam_mulai":
                    item.get(
                        "jam_mulai",
                        "",
                    ),

                "jam_selesai":
                    item.get(
                        "jam_selesai",
                        "",
                    ),

                "rentang_jam":
                    item.get(
                        "rentang_jam",
                        "",
                    ),

                "total":
                    item.get(
                        "total",
                        "",
                    ),

                "total_menit":
                    int(
                        item.get(
                            "total_menit",
                            0,
                        )
                        or 0
                    ),

                "laporan":
                    item.get(
                        "laporan",
                        "",
                    ),
            })

        first_item = (
            template_items[0]
        )

        values = {

            "NAMA":
                nama,

            "NIP":
                nip,

            "JABATAN":
                jabatan,

            "UNIT_KERJA":
                unit_kerja,

            "TANGGAL":
                first_item[
                    "tanggal"
                ],

            "TANGGAL_LEMBUR":
                first_item[
                    "tanggal"
                ],

            "TANGGAL_ISO":
                first_item[
                    "tanggal_iso"
                ],

            "TANGGAL_SURAT":
                tanggal_surat_text,

            "TANGGAL_PEMBUATAN_SURAT":
                tanggal_surat_text,

            "TANGGAL_SURAT_ISO":
                tanggal_surat_iso,

            "JAM_MULAI":
                first_item[
                    "jam_mulai"
                ],

            "JAM_SELESAI":
                first_item[
                    "jam_selesai"
                ],

            "RENTANG_JAM":
                first_item[
                    "rentang_jam"
                ],

            "RENTANG_JAM_LEMBUR":
                first_item[
                    "rentang_jam"
                ],

            "TOTAL_JAM":
                first_item[
                    "total"
                ],

            "TOTAL_JAM_LEMBUR":
                first_item[
                    "total"
                ],

            "LAPORAN_LEMBUR":
                first_item[
                    "laporan"
                ],

            "LEMBUR_ITEMS":
                template_items,
        }

        employee = {

            "id":
                employee_id,

            "nama":
                nama,

            "nip":
                nip,

            "jabatan":
                jabatan,

            "unit_kerja":
                unit_kerja,
        }

        template_name = (
            "DAFTAR HADIR KERJA LEMBUR.docx"
        )

        self.start_lembur_worker(
            template_name,
            employee,
            tanggal_surat_text,
            tanggal_surat_iso,
            values,
            employee_id,
            template_items,
        )

    # ========================================================
    # START LEMBUR WORKER
    # ========================================================

    def start_lembur_worker(
        self,
        template_name,
        employee,
        tanggal_surat_text,
        tanggal_surat_iso,
        values,
        employee_id,
        items,
    ):

        timestamp = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )

        safe_name = self.safe_filename(
            employee.get(
                "nama"
            )
            or "pegawai"
        )

        output_name = (
            f"Lembur_"
            f"{safe_name}_"
            f"{timestamp}.docx"
        )

        self.set_save_buttons_enabled(
            False
        )

        self.progress = QProgressDialog(
            "Sedang membuat surat lembur...",
            None,
            0,
            0,
            self,
        )

        self.progress.setWindowTitle(
            "LetterSync"
        )

        self.progress.setCancelButton(
            None
        )

        self.progress.setMinimumDuration(
            0
        )

        self.progress.setWindowModality(
            Qt.WindowModality.WindowModal
        )

        self.progress.show()

        self.statusBar().showMessage(
            "Sedang membuat surat lembur..."
        )

        QApplication.processEvents()

        def do_work():

            # ------------------------------------------------
            # WORD
            # ------------------------------------------------

            render_template(
                template_name,
                output_name,
                values,
            )

            docx_path = (
                WORD_DIR
                / output_name
            )

            if not docx_path.exists():

                raise FileNotFoundError(
                    "File Word tidak ditemukan:\n"
                    f"{docx_path}"
                )

            # ------------------------------------------------
            # PDF
            # ------------------------------------------------

            pdf_path = create_pdf_from_docx(
                docx_path
            )

            # ------------------------------------------------
            # DATABASE
            # ------------------------------------------------

            conn = get_connection()

            try:

                for item in items:

                    conn.execute("""
                        INSERT INTO lembur (
                            pegawai_id,
                            tanggal,
                            tanggal_surat,
                            jam_mulai,
                            jam_selesai,
                            rentang_jam,
                            total_jam,
                            laporan_lembur,
                            status,
                            file_path,
                            pdf_path,
                            created_at
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        employee_id,

                        item.get(
                            "tanggal_iso",
                            "",
                        ),

                        tanggal_surat_iso,

                        item.get(
                            "jam_mulai",
                            "",
                        ),

                        item.get(
                            "jam_selesai",
                            "",
                        ),

                        item.get(
                            "rentang_jam",
                            "",
                        ),

                        item.get(
                            "total",
                            "",
                        ),

                        item.get(
                            "laporan",
                            "",
                        ),

                        "Selesai",

                        str(
                            docx_path
                        ),

                        (
                            str(
                                pdf_path
                            )
                            if pdf_path
                            else None
                        ),

                        datetime.now().isoformat(
                            timespec="seconds"
                        ),
                    ))

                conn.commit()

            except Exception:

                conn.rollback()

                raise

            finally:

                conn.close()

            return {

                "word":
                    str(
                        docx_path
                    ),

                "pdf":
                    (
                        str(
                            pdf_path
                        )
                        if pdf_path
                        else ""
                    ),

                "jenis":
                    "Lembur",

                "jumlah_tanggal":
                    len(items),
            }

        self.worker_thread = QThread()

        self.worker = LetterWorker(
            do_work
        )

        self.worker.moveToThread(
            self.worker_thread
        )

        self.worker_thread.started.connect(
            self.worker.run
        )

        self.worker.finished.connect(
            self.lembur_finished
        )

        self.worker.error.connect(
            self.lembur_error
        )

        self.worker.finished.connect(
            self.worker_thread.quit
        )

        self.worker.error.connect(
            self.worker_thread.quit
        )

        self.worker_thread.finished.connect(
            self.worker.deleteLater
        )

        self.worker_thread.finished.connect(
            self.worker_thread.deleteLater
        )

        self.worker_thread.start()

    # ========================================================
    # LEMBUR SUCCESS
    # ========================================================

    def lembur_finished(
        self,
        result,
    ):

        self.set_save_buttons_enabled(
            True
        )

        if (
            hasattr(
                self,
                "progress",
            )
            and self.progress
        ):

            self.progress.close()

            self.progress.deleteLater()

            self.progress = None

        self.statusBar().showMessage(
            "Surat lembur berhasil dibuat."
        )

        try:

            self.load_archive()

        except Exception as e:

            print(
                "Gagal refresh Arsip:",
                e,
            )

        word_path = result.get(
            "word",
            "",
        )

        pdf_path = result.get(
            "pdf",
            "",
        )

        jumlah = result.get(
            "jumlah_tanggal",
            0,
        )

        msg = QMessageBox(
            self
        )

        msg.setWindowTitle(
            "Berhasil"
        )

        msg.setIcon(
            QMessageBox.Icon.Information
        )

        msg.setText(
            "Surat lembur berhasil dibuat."
        )

        msg.setInformativeText(
            f"{jumlah} tanggal lembur "
            "tersimpan dalam satu dokumen."
        )

        open_word = msg.addButton(
            "Buka Word",
            QMessageBox.ButtonRole.AcceptRole,
        )

        open_pdf = msg.addButton(
            "Buka PDF",
            QMessageBox.ButtonRole.AcceptRole,
        )

        open_folder = msg.addButton(
            "Buka Folder",
            QMessageBox.ButtonRole.AcceptRole,
        )

        msg.addButton(
            "Tutup",
            QMessageBox.ButtonRole.RejectRole,
        )

        msg.exec()

        clicked = msg.clickedButton()

        if clicked == open_word:

            if word_path:

                QDesktopServices.openUrl(
                    QUrl.fromLocalFile(
                        word_path
                    )
                )

        elif clicked == open_pdf:

            if pdf_path:

                QDesktopServices.openUrl(
                    QUrl.fromLocalFile(
                        pdf_path
                    )
                )

        elif clicked == open_folder:

            if word_path:

                QDesktopServices.openUrl(
                    QUrl.fromLocalFile(
                        str(
                            Path(
                                word_path
                            ).parent
                        )
                    )
                )

        self.lembur_report.clear()

    # ========================================================
    # LEMBUR ERROR
    # ========================================================

    def lembur_error(
        self,
        error_text,
    ):

        self.set_save_buttons_enabled(
            True
        )

        if (
            hasattr(
                self,
                "progress",
            )
            and self.progress
        ):

            self.progress.close()

            self.progress.deleteLater()

            self.progress = None

        self.statusBar().showMessage(
            "Gagal membuat surat lembur."
        )

        QMessageBox.critical(
            self,
            "Gagal Membuat Surat Lembur",
            (
                "Terjadi kesalahan saat membuat "
                "surat lembur.\n\n"
                + error_text
            ),
        )

    # ========================================================
    # ARSIP
    # ========================================================

    def create_archive_tab(self):

        tab = QWidget()

        layout = QVBoxLayout(
            tab
        )

        buttons = QHBoxLayout()

        refresh = QPushButton(
            "Refresh Arsip"
        )

        refresh.clicked.connect(
            self.load_archive
        )

        select_all_btn = QPushButton(
            "Pilih Semua"
        )

        select_all_btn.clicked.connect(
            self.select_all_archive
        )

        delete_btn = QPushButton(
            "Hapus yang Dipilih"
        )

        delete_btn.clicked.connect(
            self.delete_archive
        )

        buttons.addWidget(
            refresh
        )

        buttons.addWidget(
            select_all_btn
        )

        buttons.addWidget(
            delete_btn
        )

        buttons.addStretch()

        layout.addLayout(
            buttons
        )

        self.archive_table = QTableWidget()

        self.archive_table.setColumnCount(
            9
        )

        self.archive_table.setHorizontalHeaderLabels([
            "Pilih",
            "ID",
            "Jenis",
            "Pegawai",
            "Tanggal",
            "Status",
            "Word",
            "PDF",
            "Dibuat",
        ])

        self.archive_table.setAlternatingRowColors(
            True
        )

        self.archive_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )

        self.archive_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )

        self.archive_table.setColumnWidth(
            0,
            60
        )

        layout.addWidget(
            self.archive_table
        )

        self.tabs.addTab(
            tab,
            "Arsip"
        )

    # ========================================================
    # LOAD ARCHIVE
    # ========================================================

    def load_archive(self):

        if not hasattr(
            self,
            "archive_table",
        ):
            return

        conn = get_connection()

        cur = conn.cursor()

        # ----------------------------------------------------
        # SURAT BIASA
        # ----------------------------------------------------

        cur.execute("""
            SELECT
                surat.id AS id,
                surat.jenis AS jenis,
                pegawai.nama AS pegawai,
                surat.tanggal_surat AS tanggal,
                surat.status AS status,
                surat.file_path AS file_path,
                surat.pdf_path AS pdf_path,
                surat.created_at AS created_at
            FROM surat
            LEFT JOIN pegawai
                ON pegawai.id = surat.pegawai_id
        """)

        surat_rows = cur.fetchall()

        # ----------------------------------------------------
        # LEMBUR
        # ----------------------------------------------------

        cur.execute("""
            SELECT
                lembur.id AS id,
                'Lembur' AS jenis,
                pegawai.nama AS pegawai,
                lembur.tanggal AS tanggal,
                lembur.tanggal_surat AS tanggal_surat,
                lembur.status AS status,
                lembur.file_path AS file_path,
                lembur.pdf_path AS pdf_path,
                lembur.created_at AS created_at
            FROM lembur
            LEFT JOIN pegawai
                ON pegawai.id = lembur.pegawai_id
        """)

        lembur_rows = cur.fetchall()

        conn.close()

        archive_rows = []

        # ----------------------------------------------------
        # SURAT BIASA
        # ----------------------------------------------------

        for row in surat_rows:

            archive_rows.append({

                "id":
                    str(
                        row["id"]
                    ),

                "jenis":
                    row["jenis"]
                    or "",

                "pegawai":
                    row["pegawai"]
                    or "",

                "tanggal":
                    row["tanggal"]
                    or "",

                "status":
                    row["status"]
                    or "",

                "file_path":
                    row["file_path"]
                    or "",

                "pdf_path":
                    row["pdf_path"]
                    or "",

                "created_at":
                    row["created_at"]
                    or "",
            })

        # ----------------------------------------------------
        # GROUP LEMBUR
        # ----------------------------------------------------

        lembur_groups = {}

        for row in lembur_rows:

            file_path = (
                row["file_path"]
                or f"__LEMBUR_ID_{row['id']}"
            )

            if file_path not in lembur_groups:

                lembur_groups[
                    file_path
                ] = {

                    "ids": [],

                    "jenis":
                        "Lembur",

                    "pegawai":
                        row["pegawai"]
                        or "",

                    "tanggal": [],

                    "status":
                        row["status"]
                        or "",

                    "file_path":
                        row["file_path"]
                        or "",

                    "pdf_path":
                        row["pdf_path"]
                        or "",

                    "created_at":
                        row["created_at"]
                        or "",
                }

            group = (
                lembur_groups[
                    file_path
                ]
            )

            group["ids"].append(
                str(
                    row["id"]
                )
            )

            tanggal = (
                row["tanggal"]
                or row["tanggal_surat"]
                or ""
            )

            if (
                tanggal
                and tanggal
                not in group["tanggal"]
            ):

                group["tanggal"].append(
                    tanggal
                )

            if not group["file_path"]:

                group["file_path"] = (
                    row["file_path"]
                    or ""
                )

            if not group["pdf_path"]:

                group["pdf_path"] = (
                    row["pdf_path"]
                    or ""
                )

            if not group["created_at"]:

                group["created_at"] = (
                    row["created_at"]
                    or ""
                )

        # ----------------------------------------------------
        # MASUKKAN GROUP
        # ----------------------------------------------------

        for group in (
            lembur_groups.values()
        ):

            group["tanggal"].sort()

            archive_rows.append({

                "id":
                    ", ".join(
                        group["ids"]
                    ),

                "jenis":
                    group["jenis"],

                "pegawai":
                    group["pegawai"],

                "tanggal":
                    ", ".join(
                        group["tanggal"]
                    ),

                "status":
                    group["status"],

                "file_path":
                    group["file_path"],

                "pdf_path":
                    group["pdf_path"],

                "created_at":
                    group["created_at"],
            })

        # ----------------------------------------------------
        # TERBARU
        # ----------------------------------------------------

        archive_rows.sort(
            key=lambda x:
            x.get(
                "created_at",
                "",
            )
            or "",
            reverse=True,
        )

        # ----------------------------------------------------
        # TAMPILKAN
        # ----------------------------------------------------

        self.archive_table.setRowCount(
            len(
                archive_rows
            )
        )

        for r, row in enumerate(
            archive_rows
        ):

            # ------------------------------------------------
            # CHECKBOX PILIH
            # ------------------------------------------------

            check_item = QTableWidgetItem()

            check_item.setCheckState(
                Qt.CheckState.Unchecked
            )

            check_item.setTextAlignment(
                Qt.AlignmentFlag.AlignCenter
            )

            self.archive_table.setItem(
                r,
                0,
                check_item
            )

            # ------------------------------------------------
            # DATA ARSIP
            # ------------------------------------------------

            values = [

                row["id"],

                row["jenis"],

                row["pegawai"],

                row["tanggal"],

                row["status"],

                row["file_path"],

                row["pdf_path"],

                row["created_at"],
            ]

            for c, value in enumerate(
                values,
                start=1
            ):

                item = QTableWidgetItem(
                    str(
                        value
                    )
                )

                self.archive_table.setItem(
                    r,
                    c,
                    item
                )

    # ========================================================
    # DELETE ARCHIVE
    # ========================================================

    def select_all_archive(self):
        """Memilih atau membatalkan pilihan semua arsip."""

        if not hasattr(
            self,
            "archive_table",
        ):
            return

        table = self.archive_table

        if table.rowCount() == 0:
            return

        all_checked = True

        for row in range(
            table.rowCount()
        ):

            item = table.item(
                row,
                0,
            )

            if (
                not item
                or item.checkState()
                != Qt.CheckState.Checked
            ):
                all_checked = False
                break

        new_state = (
            Qt.CheckState.Unchecked
            if all_checked
            else Qt.CheckState.Checked
        )

        for row in range(
            table.rowCount()
        ):

            item = table.item(
                row,
                0,
            )

            if item:
                item.setCheckState(
                    new_state
                )


    def delete_archive(self):
        """Menghapus semua arsip yang dicentang."""

        if not hasattr(
            self,
            "archive_table",
        ):
            return

        table = self.archive_table

        selected_rows = []

        for row in range(
            table.rowCount()
        ):

            check_item = table.item(
                row,
                0,
            )

            if (
                check_item
                and check_item.checkState()
                == Qt.CheckState.Checked
            ):
                selected_rows.append(
                    row
                )

        if not selected_rows:

            QMessageBox.information(
                self,
                "Hapus Arsip",
                "Silakan centang arsip yang ingin dihapus.",
            )

            return

        # ----------------------------------------------------
        # KUMPULKAN DATA ARSIP
        # ----------------------------------------------------

        records = []

        for row in selected_rows:

            id_item = table.item(
                row,
                1,
            )

            jenis_item = table.item(
                row,
                2,
            )

            pegawai_item = table.item(
                row,
                3,
            )

            word_item = table.item(
                row,
                6,
            )

            pdf_item = table.item(
                row,
                7,
            )

            if not id_item or not jenis_item:
                continue

            record_id = (
                id_item
                .text()
                .strip()
            )

            jenis = (
                jenis_item
                .text()
                .strip()
            )

            pegawai = (
                pegawai_item.text().strip()
                if pegawai_item
                else ""
            )

            word_path = (
                word_item.text().strip()
                if word_item
                else ""
            )

            pdf_path = (
                pdf_item.text().strip()
                if pdf_item
                else ""
            )

            records.append({
                "id": record_id,
                "jenis": jenis,
                "pegawai": pegawai,
                "word": word_path,
                "pdf": pdf_path,
            })

        if not records:

            QMessageBox.information(
                self,
                "Hapus Arsip",
                "Tidak ada arsip valid yang dipilih.",
            )

            return

        # ----------------------------------------------------
        # KONFIRMASI
        # ----------------------------------------------------

        jumlah = len(records)

        preview = []

        for record in records[:10]:

            preview.append(
                f"- {record['jenis']} - "
                f"{record['pegawai']}"
            )

        if jumlah > 10:

            preview.append(
                f"- ... dan {jumlah - 10} arsip lainnya"
            )

        daftar = "\n".join(
            preview
        )

        reply = QMessageBox.question(
            self,
            "Konfirmasi Hapus",
            (
                f"Apakah {jumlah} arsip yang dipilih "
                "akan dihapus?\n\n"
                f"{daftar}\n\n"
                "Data arsip serta file Word dan PDF "
                "terkait akan dihapus."
            ),
            QMessageBox.StandardButton.Yes
            | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )

        if reply != QMessageBox.StandardButton.Yes:
            return

        conn = get_connection()

        berhasil = 0

        try:

            # ------------------------------------------------
            # HAPUS DATA DATABASE
            # ------------------------------------------------

            for record in records:

                jenis = record["jenis"]

                record_id = record["id"]

                word_path = record["word"]

                if jenis.lower() == "lembur":

                    # Lembur yang mempunyai file_path
                    # merupakan satu group beberapa tanggal.
                    if word_path:

                        conn.execute(
                            """
                            DELETE FROM lembur
                            WHERE file_path = ?
                            """,
                            (
                                word_path,
                            ),
                        )

                    else:

                        # Lembur tanpa file_path:
                        # ID dapat berupa beberapa ID.
                        ids = [
                            x.strip()
                            for x in record_id.split(",")
                            if x.strip()
                        ]

                        for lembur_id in ids:

                            conn.execute(
                                """
                                DELETE FROM lembur
                                WHERE id = ?
                                """,
                                (
                                    lembur_id,
                                ),
                            )

                else:

                    conn.execute(
                        """
                        DELETE FROM surat
                        WHERE id = ?
                        """,
                        (
                            record_id,
                        ),
                    )

                berhasil += 1

            conn.commit()

            conn.close()

            # ------------------------------------------------
            # HAPUS FILE WORD DAN PDF
            # ------------------------------------------------

            for record in records:

                for file_path in [
                    record["word"],
                    record["pdf"],
                ]:

                    if not file_path:
                        continue

                    try:

                        path = Path(
                            file_path
                        )

                        if path.exists():
                            path.unlink()

                    except Exception as e:

                        print(
                            f"Gagal menghapus file "
                            f"{file_path}: {e}"
                        )

            # ------------------------------------------------
            # REFRESH ARSIP
            # ------------------------------------------------

            self.load_archive()

            self.statusBar().showMessage(
                f"{berhasil} arsip berhasil dihapus.",
                7000,
            )

            QMessageBox.information(
                self,
                "Berhasil",
                (
                    f"{berhasil} arsip dan file "
                    "terkait berhasil dihapus."
                ),
            )

        except Exception as e:

            try:
                conn.close()
            except Exception:
                pass

            QMessageBox.critical(
                self,
                "Gagal Menghapus Arsip",
                (
                    "Arsip tidak dapat dihapus.\n\n"
                    f"{e}"
                ),
            )


    # ========================================================
    # FILE OPEN
    # ========================================================

    def open_file(
        self,
        path,
    ):

        path = Path(
            path
        )

        if not path.exists():

            QMessageBox.warning(
                self,
                "File Tidak Ditemukan",
                str(
                    path
                ),
            )

            return

        try:

            os.startfile(
                str(
                    path
                )
            )

        except Exception as e:

            QMessageBox.warning(
                self,
                "Tidak Bisa Membuka File",
                str(e),
            )

    def open_folder(
        self,
        path,
    ):

        path = Path(
            path
        )

        if not path.exists():

            QMessageBox.warning(
                self,
                "Folder Tidak Ditemukan",
                str(
                    path
                ),
            )

            return

        try:

            os.startfile(
                str(
                    path
                )
            )

        except Exception as e:

            QMessageBox.warning(
                self,
                "Tidak Bisa Membuka Folder",
                str(e),
            )


# ============================================================
# MAIN
# ============================================================

def main():

    try:

        init_db()

        ensure_archive_columns()

        ensure_lembur_table()

    except Exception:

        traceback.print_exc()

        app = QApplication.instance()

        if app is None:
            app = QApplication(sys.argv)

        QMessageBox.critical(
            None,
            "Database Error",
            "Database LetterSync tidak dapat dibuka.",
        )

        return 1

    app = QApplication(
        sys.argv
    )

    app.setApplicationName(
        "LetterSync"
    )

    app.setStyle(
        "Fusion"
    )

    window = MainWindow()

    window.show()

    return app.exec()


if __name__ == "__main__":

    sys.exit(
        main()
    )

