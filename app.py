"""
Log Detail & QR Generator — Sistem Inventarisasi Kayu Kehutanan
Versi: 3.3 | Framework: PyQt6
© 2026 Herlan. All rights reserved.
"""

import sys
import os
import math
import tempfile
from datetime import datetime
from io import BytesIO

import pandas as pd
import qrcode
from PIL import Image

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QFileDialog, QFrame,
    QMessageBox, QGroupBox, QGridLayout, QStatusBar,
    QComboBox,
)
from PyQt6.QtCore import Qt, QRect, QTimer, QMarginsF, QSettings
from PyQt6.QtGui import (
    QPixmap, QImage, QFont, QColor, QPainter, QPen, QPalette,
    QPageSize, QPageLayout, QIcon,
)
from PyQt6.QtPrintSupport import QPrinter


# ── Colour palette ──────────────────────────────────────────────────────────
C_PRIMARY   = "#1565C0"
C_PRIMARY_H = "#1976D2"
C_PRIMARY_D = "#0D47A1"
C_SUCCESS   = "#2E7D32"
C_SUCCESS_L = "#E8F5E9"
C_BG        = "#F0F2F7"
C_CARD      = "#FFFFFF"
C_BORDER    = "#DDE1E7"
C_TEXT      = "#1A1A2E"
C_TEXT_MUTE = "#607D8B"
C_ROW_ALT   = "#F0F4FF"
C_COPY      = "#6A1B9A"

ALL_ITEM = "— Semua —"

# ── Column aliases ──────────────────────────────────────────────────────────
COL_ALIASES: dict[str, list[str]] = {
    "petak":            ["petak", "blok", "block"],
    "no_batang":        ["no. batang", "no batang", "nobatang", "no_batang",
                         "nomor batang", "no.batang"],
    "id_barcode":       ["id barcode", "id_barcode", "barcode", "idbarcode",
                         "id", "kode batang"],
    "jenis_kayu":       ["jenis kayu", "jenis_kayu", "jeniskayu", "species", "kayu"],
    "panjang":          ["panjang", "panjang (m)", "length", "p (m)", "p(m)"],
    "diameter_pangkal": ["diameter pangkal", "d pangkal", "dpangkal",
                         "d. pangkal", "dp", "pangkal (cm)", "d_pangkal"],
    "diameter_ujung":   ["diameter ujung", "d ujung", "dujung", "d. ujung",
                         "du", "ujung (cm)", "d_ujung"],
    "diameter_rata":    ["diameter rata-rata", "diameter rata rata", "d rata",
                         "drata", "d_rata", "rata-rata", "rata rata"],
    "persen_cacat":     ["persen cacat", "% cacat", "cacat (%)",
                         "persen_cacat", "defect %"],
    "jenis_cacat":      ["jenis cacat", "jenis_cacat", "tipe cacat", "defect type"],
    "volume":           ["volume", "volume (m3)", "volume (m³)", "vol",
                         "vol (m3)", "m3"],
}


# ── Pure helpers ────────────────────────────────────────────────────────────
def _find_col(df: pd.DataFrame, key: str) -> str | None:
    lower_map = {c.strip().lower(): c for c in df.columns}
    for alias in COL_ALIASES.get(key, []):
        if alias in lower_map:
            return lower_map[alias]
    return None


def _val(row, col_name):
    if col_name is None:
        return None
    v = row.get(col_name)
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    return v


def floor_int(value) -> int | None:
    if value is None:
        return None
    try:
        return math.floor(float(value))
    except Exception:
        return None


def calc_volume(d_rata_cm, panjang_m, persen_cacat=0) -> float | None:
    """
    Volume = PI/4 * (d_rata/100)^2 * panjang - volume_gerowong
    volume_gerowong = PI/4 * (d_rata/100)^2 * panjang * persen_cacat/100
    Simplifikasi: volume = PI/4 * (d_rata/100)^2 * panjang * (1 - persen_cacat/100)
    """
    try:
        pct = float(persen_cacat) if persen_cacat is not None else 0.0
        d_m = float(d_rata_cm) / 100.0
        vol = (math.pi / 4.0) * d_m ** 2 * float(panjang_m)
        return round(vol * (1.0 - pct / 100.0), 2)
    except Exception:
        return None


def _setup_printer_page(printer: QPrinter):
    layout = QPageLayout(
        QPageSize(QPageSize.PageSizeId.A5),
        QPageLayout.Orientation.Portrait,
        QMarginsF(0, 0, 0, 0),
    )
    printer.setPageLayout(layout)


def _resource_path(rel: str) -> str:
    """Resolve path whether running from source or PyInstaller bundle."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def _app_icon() -> QIcon:
    path = _resource_path("icon.png")
    if os.path.exists(path):
        return QIcon(path)
    return QIcon()


def _render_to_pdf(canvas, path: str) -> bool:
    printer = QPrinter(QPrinter.PrinterMode.ScreenResolution)
    printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
    printer.setOutputFileName(path)
    _setup_printer_page(printer)
    painter = QPainter()
    if not painter.begin(printer):
        return False
    try:
        canvas.render(painter, painter.viewport())
    finally:
        painter.end()
    return True


def make_qr_pixmap(data: str, size: int = 320) -> QPixmap:
    qr = qrcode.QRCode(version=None,
                       error_correction=qrcode.constants.ERROR_CORRECT_M,
                       box_size=8, border=2)
    qr.add_data(data)
    qr.make(fit=True)
    img: Image.Image = qr.make_image(fill_color="black", back_color="white")
    buf = BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    qi = QImage()
    qi.loadFromData(buf.read())
    return QPixmap.fromImage(qi).scaled(
        size, size,
        Qt.AspectRatioMode.KeepAspectRatio,
        Qt.TransformationMode.SmoothTransformation,
    )


# ── ValueLabel ───────────────────────────────────────────────────────────────
class ValueLabel(QLabel):
    def __init__(self, accent=False, mono=False, parent=None):
        super().__init__("—", parent)
        self.setMinimumHeight(42)
        self.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)
        self.setFont(QFont("Consolas" if mono else "Segoe UI", 12))
        if accent:
            self.setStyleSheet(f"""
                QLabel {{ background:{C_SUCCESS_L}; border:1px solid #A5D6A7;
                border-radius:6px; padding:6px 14px; color:{C_SUCCESS};
                font-weight:bold; font-size:15pt; }}""")
        else:
            self.setStyleSheet(f"""
                QLabel {{ background:{C_CARD}; border:1px solid {C_BORDER};
                border-radius:6px; padding:6px 14px; color:{C_TEXT};
                font-size:12pt; }}""")

    def set_value(self, val, decimals=None, suffix=""):
        if val is None:
            self.setText("—")
            return
        self.setText(f"{float(val):.{decimals}f}{suffix}" if decimals is not None
                     else f"{val}{suffix}")


# ── PrintCanvas ──────────────────────────────────────────────────────────────
class PrintCanvas:
    def __init__(self, fields: dict, qr_data: str, sheet_name: str = ""):
        self.fields     = fields
        self.qr_data    = qr_data
        self.sheet_name = sheet_name

    def render(self, painter: QPainter, page_rect: QRect):
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setRenderHint(QPainter.RenderHint.TextAntialiasing)

        M = 36
        x, y = page_rect.x() + M, page_rect.y() + M
        W = page_rect.width() - 2 * M

        # Header bar
        painter.fillRect(x, y, W, 52, QColor(C_PRIMARY))
        painter.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        painter.setPen(QColor("white"))
        painter.drawText(x + 12, y + 34, "LOG DETAIL KEHUTANAN")
        painter.setFont(QFont("Segoe UI", 7))
        painter.drawText(x + W - 240, y + 22, f"Sheet : {self.sheet_name}")
        painter.drawText(x + W - 240, y + 36,
                         f"Dicetak: {datetime.now().strftime('%d %B %Y  %H:%M')}")
        y += 62

        # QR
        qr_size = 200
        if self.qr_data and self.qr_data != "—":
            try:
                pix = make_qr_pixmap(self.qr_data, qr_size)
                painter.drawPixmap(x, y, pix)
            except Exception:
                pass
        painter.setFont(QFont("Consolas", 8))
        painter.setPen(QColor(C_TEXT))
        painter.drawText(QRect(x, y + qr_size + 4, qr_size, 20),
                         Qt.AlignmentFlag.AlignCenter, self.qr_data or "—")

        # Data fields
        dx, dy, dw, lbl_w, row_h = x + qr_size + 24, y, W - qr_size - 24, 190, 26
        field_rows = [
            ("Petak",                  "petak"),
            ("No. Batang",             "no_batang"),
            ("ID Barcode",             "id_barcode"),
            ("Jenis Kayu",             "jenis_kayu"),
            ("Panjang (m)",            "panjang"),
            ("Diameter Pangkal (cm)",  "diameter_pangkal"),
            ("Diameter Ujung (cm)",    "diameter_ujung"),
            ("Diameter Rata-Rata (cm)","diameter_rata"),
            ("Persen Cacat (%)",       "persen_cacat"),
            ("Jenis Cacat",            "jenis_cacat"),
            ("Volume (M³)",            "volume"),
        ]
        for i, (label, key) in enumerate(field_rows):
            ry = dy + i * row_h
            is_v = key == "volume"
            bg = QColor(C_SUCCESS_L) if is_v else (
                 QColor(C_ROW_ALT) if i % 2 == 0 else QColor(C_CARD))
            painter.fillRect(dx - 4, ry, dw + 4, row_h - 1, bg)
            painter.setFont(QFont("Segoe UI", 8, QFont.Weight.Bold))
            painter.setPen(QColor(C_TEXT_MUTE))
            painter.drawText(dx, ry + row_h - 7, f"{label}:")
            painter.setFont(QFont("Consolas" if key == "id_barcode" else "Segoe UI",
                                  10 if is_v else 9,
                                  QFont.Weight.Bold if is_v else QFont.Weight.Normal))
            painter.setPen(QColor(C_SUCCESS if is_v else C_TEXT))
            painter.drawText(dx + lbl_w, ry + row_h - 7, self.fields.get(key, "—"))

        # Footer
        fy = dy + len(field_rows) * row_h + 18
        painter.setPen(QPen(QColor(C_BORDER), 1))
        painter.drawLine(x, fy, x + W, fy)
        painter.setFont(QFont("Segoe UI", 7))
        painter.setPen(QColor(C_TEXT_MUTE))
        painter.drawText(x, fy + 14,
                         "Sistem Inventarisasi Kayu Kehutanan  |  Log Detail & QR Generator")
        painter.setPen(QColor(C_COPY))
        painter.setFont(QFont("Segoe UI", 7, QFont.Weight.Bold))
        painter.drawText(x + W - 190, fy + 14, "© 2026 Herlan. All rights reserved.")


# ── Main Window ──────────────────────────────────────────────────────────────
class ForestryApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.df: pd.DataFrame | None = None
        self.col_map: dict[str, str | None] = {}
        self.visible_rows: list[int] = []
        self.current_pos: int = 0
        self._xl_path: str = ""
        self._xl_sheets: list[str] = []
        self._current_sheet: str = ""
        self._loading: bool = False
        self._settings = QSettings("Herlan", "LogDetailQR")

        self.setWindowTitle("Log Detail & QR Generator  —  Sistem Kehutanan")
        self.setWindowIcon(_app_icon())
        self.setMinimumSize(1200, 800)
        self.resize(1440, 900)

        self._build_ui()
        self._apply_styles()
        QTimer.singleShot(100, self._restore_last_file)

    # ── UI ────────────────────────────────────────────────────────────────
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setSpacing(8)
        root.setContentsMargins(14, 14, 14, 8)

        # ── ROW 1: File + Sheet ──────────────────────────────────────
        r1 = QHBoxLayout()
        r1.setSpacing(8)

        fg = QGroupBox("Sumber Data Excel")
        fl = QHBoxLayout(fg)
        fl.setContentsMargins(10, 6, 10, 6)
        self.lbl_file = QLabel("Belum ada file dipilih")
        self.lbl_file.setStyleSheet(f"color:{C_TEXT_MUTE}; font-style:italic; font-size:11pt;")
        self.btn_open = QPushButton("Pilih File .xlsx")
        self.btn_open.setFixedHeight(36)
        self.btn_open.clicked.connect(self._open_file)
        fl.addWidget(self.lbl_file, 1)
        fl.addWidget(self.btn_open)
        r1.addWidget(fg, 3)

        shg = QGroupBox("Sheet")
        shl = QHBoxLayout(shg)
        shl.setContentsMargins(10, 6, 10, 6)
        shl.setSpacing(8)
        lbl_sh = QLabel("Pilih Sheet:")
        lbl_sh.setStyleSheet(f"color:{C_TEXT_MUTE}; font-weight:bold; font-size:11pt;")
        self.cmb_sheet = QComboBox()
        self.cmb_sheet.setFixedHeight(34)
        self.cmb_sheet.setEnabled(False)
        self.cmb_sheet.setPlaceholderText("— belum ada file —")
        self.cmb_sheet.currentIndexChanged.connect(self._on_sheet_index_changed)
        shl.addWidget(lbl_sh)
        shl.addWidget(self.cmb_sheet, 1)
        r1.addWidget(shg, 2)

        root.addLayout(r1)

        # ── ROW 2: Filter ────────────────────────────────────────────
        r2 = QGroupBox("Filter Kolom")
        r2l = QHBoxLayout(r2)
        r2l.setContentsMargins(12, 6, 12, 6)
        r2l.setSpacing(12)

        self._filter_combos: dict[str, QComboBox] = {}

        for key, label in [("petak", "Petak"), ("jenis_kayu", "Jenis Kayu"),
                            ("jenis_cacat", "Jenis Cacat")]:
            lbl = QLabel(f"{label}:")
            lbl.setStyleSheet(f"font-weight:bold; color:{C_TEXT_MUTE}; font-size:10pt;")
            cmb = QComboBox()
            cmb.setFixedHeight(34)
            cmb.setMinimumWidth(150)
            cmb.setEnabled(False)
            cmb.addItem(ALL_ITEM)
            cmb.currentIndexChanged.connect(self._apply_filters)
            self._filter_combos[key] = cmb
            r2l.addWidget(lbl)
            r2l.addWidget(cmb)

        r2l.addSpacing(20)

        lbl_cari = QLabel("Cari:")
        lbl_cari.setStyleSheet(f"font-weight:bold; color:{C_TEXT_MUTE}; font-size:10pt;")
        self.cmb_search_field = QComboBox()
        self.cmb_search_field.addItems(["ID Barcode", "No. Batang", "Semua Kolom"])
        self.cmb_search_field.setFixedHeight(34)
        self.cmb_search_field.setFixedWidth(140)
        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("Ketik lalu Enter …")
        self.txt_search.setFixedHeight(34)
        self.txt_search.returnPressed.connect(self._apply_filters)
        self.txt_search.textChanged.connect(self._on_search_text_changed)
        self.btn_reset_all = QPushButton("Reset Semua Filter")
        self.btn_reset_all.setFixedHeight(34)
        self.btn_reset_all.setObjectName("btnSecondary")
        self.btn_reset_all.clicked.connect(self._reset_all_filters)

        r2l.addWidget(lbl_cari)
        r2l.addWidget(self.cmb_search_field)
        r2l.addWidget(self.txt_search, 1)
        r2l.addWidget(self.btn_reset_all)

        root.addWidget(r2)

        # ── ROW 3: Summary info bar ──────────────────────────────────
        info_frame = QFrame()
        info_frame.setObjectName("card")
        info_frame.setFixedHeight(40)
        info_l = QHBoxLayout(info_frame)
        info_l.setContentsMargins(16, 4, 16, 4)

        self.lbl_record_info = QLabel("— baris")
        self.lbl_record_info.setStyleSheet(
            f"color:{C_TEXT_MUTE}; font-size:10pt; font-weight:bold;")

        sep_v = QFrame()
        sep_v.setFrameShape(QFrame.Shape.VLine)
        sep_v.setStyleSheet(f"color:{C_BORDER};")

        self.lbl_vol_total = QLabel("Total Volume: — m³")
        self.lbl_vol_total.setStyleSheet(
            f"color:{C_SUCCESS}; font-size:10pt; font-weight:bold;")

        info_l.addWidget(self.lbl_record_info)
        info_l.addWidget(sep_v)
        info_l.addWidget(self.lbl_vol_total)
        info_l.addStretch()

        root.addWidget(info_frame)

        # ── MAIN CONTENT: QR + Detail ────────────────────────────────
        main_widget = QWidget()
        main_layout = QHBoxLayout(main_widget)
        main_layout.setSpacing(14)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # ---- QR panel ----
        qr_frame = QFrame()
        qr_frame.setObjectName("card")
        qr_frame.setFixedWidth(420)
        qrl = QVBoxLayout(qr_frame)
        qrl.setContentsMargins(18, 16, 18, 16)
        qrl.setSpacing(10)
        qrl.setAlignment(Qt.AlignmentFlag.AlignHCenter)

        qr_title = QLabel("QR Code")
        qr_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        qr_title.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        qr_title.setStyleSheet(f"color:{C_PRIMARY};")

        self.qr_display = QLabel()
        self.qr_display.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.qr_display.setFixedSize(370, 370)
        self.qr_display.setStyleSheet(f"""
            QLabel {{ border:2px dashed {C_BORDER}; border-radius:10px; background:white; }}""")
        self._clear_qr()

        self.lbl_qr_id = QLabel("—")
        self.lbl_qr_id.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_qr_id.setFont(QFont("Consolas", 11))
        self.lbl_qr_id.setWordWrap(True)
        self.lbl_qr_id.setStyleSheet(f"color:{C_TEXT_MUTE};")

        self.lbl_badge = QLabel("— / —")
        self.lbl_badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_badge.setStyleSheet(f"""
            background:{C_PRIMARY}; color:white; border-radius:14px;
            padding:6px 20px; font-weight:bold; font-size:13pt;""")

        qrl.addStretch()
        qrl.addWidget(qr_title)
        qrl.addWidget(self.qr_display, 0, Qt.AlignmentFlag.AlignHCenter)
        qrl.addWidget(self.lbl_qr_id)
        qrl.addSpacing(6)
        qrl.addWidget(self.lbl_badge, 0, Qt.AlignmentFlag.AlignHCenter)
        qrl.addStretch()
        main_layout.addWidget(qr_frame)

        # ---- Detail panel ----
        detail_frame = QFrame()
        detail_frame.setObjectName("card")
        dl = QVBoxLayout(detail_frame)
        dl.setContentsMargins(24, 14, 24, 14)
        dl.setSpacing(10)

        dtitle = QLabel("Detail Data Log Kayu")
        dtitle.setFont(QFont("Segoe UI", 15, QFont.Weight.Bold))
        dtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        dtitle.setStyleSheet(f"color:{C_PRIMARY};")
        dl.addWidget(dtitle)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet(f"color:{C_BORDER};")
        dl.addWidget(sep)

        grid = QGridLayout()
        grid.setSpacing(10)
        grid.setColumnStretch(1, 1)
        grid.setColumnStretch(3, 1)

        self._fields: dict[str, ValueLabel] = {}

        def add_field(label_text, key, row, col, accent=False, mono=False):
            lbl = QLabel(label_text)
            lbl.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))
            lbl.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            lbl.setStyleSheet(f"color:{C_TEXT_MUTE};")
            val = ValueLabel(accent=accent, mono=mono)
            self._fields[key] = val
            grid.addWidget(lbl, row, col, Qt.AlignmentFlag.AlignRight)
            grid.addWidget(val, row, col + 1)

        add_field("Petak",                   "petak",            0, 0)
        add_field("No. Batang",              "no_batang",        0, 2)
        add_field("ID Barcode",              "id_barcode",       1, 0, mono=True)
        add_field("Jenis Kayu",              "jenis_kayu",       1, 2)
        add_field("Panjang (m)",             "panjang",          2, 0)
        add_field("Jenis Cacat",             "jenis_cacat",      2, 2)
        add_field("Diameter Pangkal (cm)",   "diameter_pangkal", 3, 0)
        add_field("Persen Cacat (%)",        "persen_cacat",     3, 2)
        add_field("Diameter Ujung (cm)",     "diameter_ujung",   4, 0)
        add_field("Diameter Rata-Rata (cm)", "diameter_rata",    4, 2)

        vol_lbl = QLabel("Volume (M³)")
        vol_lbl.setFont(QFont("Segoe UI", 13, QFont.Weight.Bold))
        vol_lbl.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        vol_lbl.setStyleSheet(f"color:{C_SUCCESS};")
        vol_val = ValueLabel(accent=True)
        self._fields["volume"] = vol_val
        grid.addWidget(vol_lbl, 5, 0, Qt.AlignmentFlag.AlignRight)
        grid.addWidget(vol_val, 5, 1, 1, 3)

        dl.addLayout(grid)
        dl.addStretch()
        main_layout.addWidget(detail_frame, 1)

        root.addWidget(main_widget, 1)

        # ── BOTTOM BAR ───────────────────────────────────────────────
        bot = QHBoxLayout()
        bot.setSpacing(10)

        ng = QGroupBox("Navigasi Record")
        nl = QHBoxLayout(ng)
        nl.setContentsMargins(10, 6, 10, 6)
        nl.setSpacing(6)

        self.btn_first = QPushButton("|◀")
        self.btn_prev  = QPushButton("◀ Prev")
        self.btn_next  = QPushButton("Next ▶")
        self.btn_last  = QPushButton("▶|")

        for b in (self.btn_first, self.btn_prev, self.btn_next, self.btn_last):
            b.setFixedHeight(36)
            b.setEnabled(False)

        self.btn_first.clicked.connect(lambda: self._navigate("first"))
        self.btn_prev.clicked.connect(lambda:  self._navigate("prev"))
        self.btn_next.clicked.connect(lambda:  self._navigate("next"))
        self.btn_last.clicked.connect(lambda:  self._navigate("last"))

        nl.addWidget(self.btn_first)
        nl.addWidget(self.btn_prev)
        nl.addSpacing(4)
        nl.addWidget(self.btn_next)
        nl.addWidget(self.btn_last)
        bot.addWidget(ng, 2)

        pg = QGroupBox("Cetak / Ekspor")
        pl = QHBoxLayout(pg)
        pl.setContentsMargins(10, 6, 10, 6)
        pl.setSpacing(8)

        self.btn_preview = QPushButton("Pratinjau & Cetak")
        self.btn_preview.setFixedHeight(36)
        self.btn_preview.setEnabled(False)
        self.btn_preview.clicked.connect(self._print_preview)

        self.btn_pdf = QPushButton("Simpan PDF")
        self.btn_pdf.setFixedHeight(36)
        self.btn_pdf.setObjectName("btnSecondary")
        self.btn_pdf.setEnabled(False)
        self.btn_pdf.clicked.connect(self._save_pdf)

        pl.addWidget(self.btn_preview)
        pl.addWidget(self.btn_pdf)
        bot.addWidget(pg, 2)

        copy_lbl = QLabel("© 2026 Herlan — All rights reserved.")
        copy_lbl.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        copy_lbl.setStyleSheet(f"color:{C_COPY}; font-size:9pt; font-weight:bold;")
        bot.addWidget(copy_lbl)

        root.addLayout(bot)

        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.showMessage("Silakan pilih file Excel (.xlsx) untuk memulai.")

        self._search_timer = QTimer()
        self._search_timer.setSingleShot(True)
        self._search_timer.setInterval(350)
        self._search_timer.timeout.connect(self._apply_filters)

    # ── Styles ────────────────────────────────────────────────────────────
    def _apply_styles(self):
        self.setStyleSheet(f"""
            QMainWindow, QWidget {{
                background:{C_BG}; font-family:"Segoe UI"; font-size:10pt;
            }}
            QGroupBox {{
                font-weight:bold; font-size:10pt;
                border:1px solid {C_BORDER}; border-radius:8px;
                margin-top:12px; padding-top:6px;
                background:{C_CARD}; color:{C_TEXT};
            }}
            QGroupBox::title {{
                subcontrol-origin:margin; left:12px; padding:0 6px; color:{C_PRIMARY};
            }}
            QPushButton {{
                background:{C_PRIMARY}; color:white; border:none;
                border-radius:6px; padding:6px 18px;
                font-weight:bold; font-size:10pt;
            }}
            QPushButton:hover   {{ background:{C_PRIMARY_H}; }}
            QPushButton:pressed {{ background:{C_PRIMARY_D}; }}
            QPushButton:disabled {{ background:#CFD8DC; color:#90A4AE; }}
            QPushButton#btnSecondary {{
                background:{C_CARD}; color:{C_PRIMARY};
                border:1.5px solid {C_PRIMARY};
            }}
            QPushButton#btnSecondary:hover   {{ background:#E3F2FD; }}
            QPushButton#btnSecondary:disabled {{
                background:{C_CARD}; color:#B0BEC5; border-color:#B0BEC5;
            }}
            QLineEdit, QComboBox {{
                border:1.5px solid {C_BORDER}; border-radius:5px;
                padding:4px 10px; background:white;
                font-size:10pt; color:{C_TEXT};
            }}
            QLineEdit:focus {{ border-color:{C_PRIMARY}; }}
            QComboBox::drop-down {{ border:none; width:20px; }}
            QComboBox:disabled {{ background:#F5F5F5; color:#aaa; }}
            QFrame#card {{
                background:{C_CARD}; border:1px solid {C_BORDER};
                border-radius:12px;
            }}
            QStatusBar {{
                background:{C_CARD}; border-top:1px solid {C_BORDER};
                font-size:9pt; color:{C_TEXT_MUTE};
            }}
        """)

    # ── Session restore ───────────────────────────────────────────────────
    def _restore_last_file(self):
        last_path  = self._settings.value("last_file", "")
        last_sheet = self._settings.value("last_sheet", "")
        if not last_path or not os.path.exists(last_path):
            return
        try:
            xl = pd.ExcelFile(last_path, engine="openpyxl")
            self._xl_path   = last_path
            self._xl_sheets = xl.sheet_names
            self._xl_file   = xl

            self._loading = True
            self.cmb_sheet.clear()
            self.cmb_sheet.addItems(self._xl_sheets)
            self.cmb_sheet.setEnabled(True)
            self._loading = False

            target = last_sheet if last_sheet in self._xl_sheets else self._xl_sheets[0]
            idx = self._xl_sheets.index(target)
            self.cmb_sheet.setCurrentIndex(idx)
            self._load_sheet(target)
        except Exception:
            pass  # Jika file rusak / sudah dipindah, abaikan saja

    # ── File loading ──────────────────────────────────────────────────────
    def _open_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Pilih File Excel", "",
            "Excel Files (*.xlsx *.xls);;All Files (*)",
        )
        if not path:
            return
        try:
            xl = pd.ExcelFile(path, engine="openpyxl")
            self._xl_path   = path
            self._xl_sheets = xl.sheet_names
            self._xl_file   = xl

            self._loading = True
            self.cmb_sheet.clear()
            self.cmb_sheet.addItems(self._xl_sheets)
            self.cmb_sheet.setEnabled(True)
            self._loading = False

            self._settings.setValue("last_file", path)
            self.cmb_sheet.setCurrentIndex(0)
            self._load_sheet(self._xl_sheets[0])

        except Exception as exc:
            QMessageBox.critical(self, "Gagal Membaca File",
                                 f"Error:\n{exc}\n\nPastikan file tidak dibuka di Excel.")

    def _on_sheet_index_changed(self, index: int):
        if self._loading or index < 0 or index >= len(self._xl_sheets):
            return
        self._load_sheet(self._xl_sheets[index])

    def _load_sheet(self, sheet_name: str):
        try:
            df = self._xl_file.parse(sheet_name)
            df.columns = df.columns.str.strip()
            self.df = df
            self._current_sheet = sheet_name
            self.col_map = {k: _find_col(df, k) for k in COL_ALIASES}

            fname = os.path.basename(self._xl_path)
            self.lbl_file.setText(
                f"  {fname}   Sheet: {sheet_name}   |   {len(df):,} baris")
            self.lbl_file.setStyleSheet(f"color:{C_SUCCESS}; font-weight:bold; font-size:11pt;")

            self._loading = True
            for key, cmb in self._filter_combos.items():
                cmb.clear()
                cmb.addItem(ALL_ITEM)
                col = self.col_map.get(key)
                if col:
                    uniq = sorted(
                        df[col].dropna().astype(str).str.strip().unique().tolist()
                    )
                    cmb.addItems(uniq)
                cmb.setEnabled(True)
            self.txt_search.clear()
            self._loading = False

            self._enable_controls(True)
            self._apply_filters()

            missing = [k for k, v in self.col_map.items() if v is None]
            msg = (f"Sheet '{sheet_name}' dimuat — {len(df):,} baris"
                   + (f"  |  Kolom tak terdeteksi: {', '.join(missing)}" if missing else
                      "  |  Semua kolom terdeteksi"))
            self.status.showMessage(msg)
            self._settings.setValue("last_sheet", sheet_name)

        except Exception as exc:
            QMessageBox.critical(self, "Gagal Memuat Sheet", str(exc))

    # ── Filter logic ──────────────────────────────────────────────────────
    def _on_search_text_changed(self, _text: str):
        self._search_timer.start()

    def _apply_filters(self):
        if self._loading or self.df is None:
            return

        df = self.df
        mask = [True] * len(df)

        for key, cmb in self._filter_combos.items():
            val = cmb.currentText()
            if val == ALL_ITEM:
                continue
            col = self.col_map.get(key)
            if col is None:
                continue
            col_series = df[col].astype(str).str.strip()
            mask = [m and (col_series.iloc[i] == val)
                    for i, m in enumerate(mask)]

        query = self.txt_search.text().strip().lower()
        if query:
            field = self.cmb_search_field.currentText()
            if field == "Semua Kolom":
                def row_matches(i):
                    return any(
                        query in str(df.iloc[i].get(c, "")).lower()
                        for c in df.columns
                    )
                mask = [m and row_matches(i) for i, m in enumerate(mask)]
            else:
                key = "id_barcode" if field == "ID Barcode" else "no_batang"
                col = self.col_map.get(key)
                if col:
                    col_series = df[col].astype(str).str.lower()
                    mask = [m and (query in col_series.iloc[i])
                            for i, m in enumerate(mask)]

        self.visible_rows = [i for i, m in enumerate(mask) if m]
        self.current_pos  = 0
        self._update_summary()
        self._display()

    def _update_summary(self):
        n = len(self.visible_rows)
        total_vol = 0.0
        for i in self.visible_rows:
            row = self.df.iloc[i]
            d_rata_raw = self._get_d_rata(row)
            d_rata_int = floor_int(d_rata_raw)
            panjang      = _val(row, self.col_map.get("panjang"))
            persen_cacat = _val(row, self.col_map.get("persen_cacat"))
            if d_rata_int is not None and panjang is not None:
                v = calc_volume(d_rata_int, panjang, persen_cacat)
                if v is not None:
                    total_vol += v
            else:
                v = _val(row, self.col_map.get("volume"))
                if v is not None:
                    try:
                        total_vol += float(v)
                    except Exception:
                        pass

        total_str = f" dari {len(self.df):,}" if n != len(self.df) else ""
        self.lbl_record_info.setText(f"{n:,} baris{total_str}")
        self.lbl_vol_total.setText(f"Total Volume: {total_vol:,.2f} m³")

    def _reset_all_filters(self):
        if self.df is None:
            return
        self._loading = True
        for cmb in self._filter_combos.values():
            cmb.setCurrentIndex(0)
        self.txt_search.clear()
        self._loading = False
        self._apply_filters()

    # ── Detail display ────────────────────────────────────────────────────
    def _get_d_rata(self, row) -> float | None:
        dp = _val(row, self.col_map.get("diameter_pangkal"))
        du = _val(row, self.col_map.get("diameter_ujung"))
        if dp is not None and du is not None:
            try:
                return (float(dp) + float(du)) / 2.0
            except Exception:
                pass
        raw = _val(row, self.col_map.get("diameter_rata"))
        if raw is not None:
            try:
                return float(raw)
            except Exception:
                pass
        return None

    def _clear_qr(self):
        ph = QPixmap(370, 370)
        ph.fill(QColor("white"))
        self.qr_display.setPixmap(ph)

    def _display(self):
        if self.df is None or not self.visible_rows:
            self._clear_qr()
            self.lbl_qr_id.setText("— Tidak ada data —")
            self.lbl_badge.setText("0 / 0")
            for v in self._fields.values():
                v.setText("—")
            return

        self.current_pos = max(0, min(self.current_pos, len(self.visible_rows) - 1))
        df_idx = self.visible_rows[self.current_pos]
        row    = self.df.iloc[df_idx]

        def g(key):
            return _val(row, self.col_map.get(key))

        petak        = g("petak")
        no_batang    = g("no_batang")
        id_barcode   = g("id_barcode")
        jenis_kayu   = g("jenis_kayu")
        panjang      = g("panjang")
        d_pangkal    = g("diameter_pangkal")
        d_ujung      = g("diameter_ujung")
        persen_cacat = g("persen_cacat")
        jenis_cacat  = g("jenis_cacat")

        d_rata_raw  = self._get_d_rata(row)
        d_rata_int  = floor_int(d_rata_raw)
        d_pangkal_int = floor_int(d_pangkal)
        d_ujung_int   = floor_int(d_ujung)

        volume = (calc_volume(d_rata_int, panjang, persen_cacat)
                  if d_rata_int is not None and panjang is not None
                  else g("volume"))

        def sv(key, val, dec=None, sfx=""):
            self._fields[key].set_value(val, dec, sfx)

        sv("petak",            petak)
        sv("no_batang",        no_batang)
        sv("id_barcode",       id_barcode)
        sv("jenis_kayu",       jenis_kayu)
        sv("panjang",          panjang,       2, " m")
        sv("diameter_pangkal", d_pangkal_int, 0, " cm")
        sv("diameter_ujung",   d_ujung_int,   0, " cm")
        sv("diameter_rata",    d_rata_int,    0, " cm")
        sv("persen_cacat",     persen_cacat,  1, " %")
        sv("jenis_cacat",      jenis_cacat)
        sv("volume",           volume,        2, " m³")

        # QR
        bc = str(id_barcode) if id_barcode is not None else ""
        if bc:
            try:
                self.qr_display.setPixmap(make_qr_pixmap(bc, 360))
                self.lbl_qr_id.setText(bc)
            except Exception as exc:
                self._clear_qr()
                self.lbl_qr_id.setText(f"QR Error: {exc}")
        else:
            self._clear_qr()
            self.lbl_qr_id.setText("— No Barcode —")

        total = len(self.visible_rows)
        self.lbl_badge.setText(f"{self.current_pos + 1:,}  /  {total:,}")
        self.status.showMessage(
            f"Record {self.current_pos + 1}/{total}  |  "
            f"Sheet: {self._current_sheet}  |  "
            f"ID: {bc or '—'}  |  Vol: {self._fields['volume'].text()}"
        )

    # ── Navigation ────────────────────────────────────────────────────────
    def _navigate(self, direction: str):
        if not self.visible_rows:
            return
        n = len(self.visible_rows)
        if   direction == "first": self.current_pos = 0
        elif direction == "prev":  self.current_pos = max(0, self.current_pos - 1)
        elif direction == "next":  self.current_pos = min(n - 1, self.current_pos + 1)
        elif direction == "last":  self.current_pos = n - 1
        self._display()

    # ── Controls enable ───────────────────────────────────────────────────
    def _enable_controls(self, on: bool):
        for w in (self.btn_first, self.btn_prev, self.btn_next, self.btn_last,
                  self.btn_preview, self.btn_pdf):
            w.setEnabled(on)

    # ── Print / PDF ───────────────────────────────────────────────────────
    def _make_canvas(self) -> PrintCanvas:
        return PrintCanvas(
            {k: v.text() for k, v in self._fields.items()},
            self.lbl_qr_id.text(),
            self._current_sheet,
        )

    def _print_preview(self):
        try:
            tmp_fd, tmp_path = tempfile.mkstemp(suffix=".pdf")
            os.close(tmp_fd)
            if not _render_to_pdf(self._make_canvas(), tmp_path):
                QMessageBox.critical(self, "Error Pratinjau",
                                     "Gagal membuat file pratinjau PDF.")
                return
            os.startfile(tmp_path)
            self.status.showMessage("Pratinjau dibuka di PDF viewer sistem.")
        except Exception as exc:
            QMessageBox.critical(self, "Gagal Pratinjau", str(exc))

    def _save_pdf(self):
        bc = self._fields["id_barcode"].text().replace("/", "-").replace("\\", "-")
        default = f"log_{bc}.pdf" if bc != "—" else "log_detail.pdf"
        path, _ = QFileDialog.getSaveFileName(
            self, "Simpan sebagai PDF", default, "PDF Files (*.pdf)")
        if not path:
            return
        if not path.lower().endswith(".pdf"):
            path += ".pdf"
        try:
            if not _render_to_pdf(self._make_canvas(), path):
                QMessageBox.critical(self, "Error PDF",
                                     "Gagal membuka output PDF.\n"
                                     "Pastikan folder tujuan dapat ditulis.")
                return
            self.status.showMessage(f"PDF disimpan: {path}")
            QMessageBox.information(self, "PDF Tersimpan",
                                    f"File berhasil disimpan:\n{path}")
        except Exception as exc:
            QMessageBox.critical(self, "Gagal Simpan PDF", str(exc))


# ── Entry point ──────────────────────────────────────────────────────────────
def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Log Detail & QR Generator")
    app.setOrganizationName("Herlan")
    app.setStyle("Fusion")
    app.setWindowIcon(_app_icon())

    pal = app.palette()
    pal.setColor(QPalette.ColorRole.Window,          QColor(C_BG))
    pal.setColor(QPalette.ColorRole.WindowText,      QColor(C_TEXT))
    pal.setColor(QPalette.ColorRole.Base,            QColor("white"))
    pal.setColor(QPalette.ColorRole.AlternateBase,   QColor(C_ROW_ALT))
    pal.setColor(QPalette.ColorRole.Highlight,       QColor(C_PRIMARY))
    pal.setColor(QPalette.ColorRole.HighlightedText, QColor("white"))
    app.setPalette(pal)

    win = ForestryApp()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
