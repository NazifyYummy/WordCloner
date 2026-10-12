# Here is the source code:
## wordcloner.py
'''
"""
================================================================================
 WordCloner — Letter Issuance System
================================================================================
"""

import argparse
import os
import sys
import tempfile
from datetime import date, datetime
from json import load, dump
from shutil import copy2
from copy import deepcopy
from pathlib import Path
from collections import Counter

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFrame, QGraphicsDropShadowEffect,
    QStackedWidget, QLineEdit, QFileDialog, QDialog, QComboBox,
    QScrollArea, QColorDialog, QCheckBox, QSystemTrayIcon, QMenu
)
from PySide6.QtCore import (
    Qt, QPoint, Signal, QSize, QRectF, QEvent, QTimer, QRegularExpression, QUrl
)
from PySide6.QtGui import (
    QIcon, QPainter, QColor, QFont, QPen, QPainterPath,
    QGuiApplication, QPalette, QRegularExpressionValidator, QAction,
    QDesktopServices,
)

try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.oxml.ns import qn
    from docx.table import Table as DocxTable
    from docx.text.paragraph import Paragraph as DocxParagraph
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False
    print("[warning] python-docx is not installed → pip install python-docx")

try:
    from openpyxl import load_workbook
    HAS_EXCEL = True
except ImportError:
    load_workbook = None
    HAS_EXCEL = False


# ═════════════════════════════════════════════════════════════════
#  Paths
# ═════════════════════════════════════════════════════════════════
from wordcloner_paths import (
    APP_ROOT, DOCUMENTS_DIR, WORDCLONER_DIR, WORDS_DIR,
    TEMPLATES_JSON, SETTINGS_JSON, migrate_legacy_data,
)

# ═════════════════════════════════════════════════════════════════
#  Settings
# ═════════════════════════════════════════════════════════════════

DEFAULT_SETTINGS = {
    "output_dir": str(Path.home() / "Desktop"),
    "placeholder_wrapper": '##',
    "quick_access": [],
    "language": "en",     
    "window_size": [720, 600],
    "window_maximized": False,
}
SETTINGS = dict(DEFAULT_SETTINGS)
_FORCE_ENGLISH_UI = False

def load_settings():
    global SETTINGS
    SETTINGS = dict(DEFAULT_SETTINGS)
    if SETTINGS_JSON.exists():
        try:
            with open(SETTINGS_JSON, "r", encoding="utf-8") as f:
                data = load(f)
            if isinstance(data, dict):
                SETTINGS.update(data)
                if not isinstance(SETTINGS.get("quick_access"), list):
                    SETTINGS["quick_access"] = []
                if SETTINGS.get("language") not in ("fa", "en"):
                    SETTINGS["language"] = "en"
        except Exception as e:
            print(f"[error] reading settings: {e}")


def save_settings():
    WORDCLONER_DIR.mkdir(parents=True, exist_ok=True)
    try:
        with open(SETTINGS_JSON, "w", encoding="utf-8") as f:
            dump(SETTINGS, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"[error] saving settings: {e}")


# ═════════════════════════════════════════════════════════════════
#  i18n
# ═════════════════════════════════════════════════════════════════
from wordcloner_translations import T

LANG_MAP = {"fa": "persian", "en": "english"}


def tr(item, **kwargs):
    if not isinstance(item, dict):
        return str(item)

    if _FORCE_ENGLISH_UI:
        value = item.get("english")
        if value is None:
            value = item.get("persian", "")
    else:
        key = LANG_MAP.get(SETTINGS.get("language", "fa"), "persian")
        value = item.get(key)
        if value is None:
            value = item.get("persian", "")

    if kwargs:
        try:
            value = value.format(**kwargs)
        except Exception:
            pass
    return value

def _is_rtl():
    return SETTINGS.get("language", "fa") == "fa"


def _dir():
    return Qt.RightToLeft if _is_rtl() else Qt.LeftToRight

# ═════════════════════════════════════════════════════════════════
#  Colors (only green theme)
# ═════════════════════════════════════════════════════════════════

C_DARK_GREEN  = "#2E5A4C"
C_GREEN       = "#3A7A5E"
C_LIGHT_GREEN = "#E8F0EC"
C_HOVER       = "#F0F5F2"
C_BG          = "#F4F7F2"
C_BORDER      = "#D0DDD5"
C_MUTED       = "#7A8F86"
C_ACCENT1     = "#8FB5A2"
C_ACCENT2     = "#E6EEE8"
C_RED         = "#E0453A"


# ═════════════════════════════════════════════════════════════════
#  Icons
# ═════════════════════════════════════════════════════════════════
from wordcloner_icons import (
    icon_document, icon_document_big, icon_plus_circle, icon_plus_white,
    icon_trash, icon_chevron_down, icon_grip, icon_upload, icon_arrow_left,
    icon_gear, icon_help, icon_status, icon_minimize, icon_maximize, icon_close,
    icon_check,
)


def apply_color_button_style(btn, hex_color):
    btn.setStyleSheet(f"""
        QPushButton {{
            background-color: {hex_color};
            border: 2px solid {C_GREEN};
            border-radius: 8px;
        }}
        QPushButton:hover {{
            border: 2px solid {C_DARK_GREEN};
        }}
    """)

# ═════════════════════════════════════════════════════════════════
#  ComboWithArrow
# ═════════════════════════════════════════════════════════════════

class ComboWithArrow(QComboBox):
    ARROW_AREA = 24

    _ARROW_QSS = """
        QComboBox::drop-down {
            border: none;
            width: 0px;
            background: transparent;
        }
        QComboBox::down-arrow {
            image: none;
            width: 0px;
            height: 0px;
        }
    """

    def __init__(self, editable=False, parent=None):
        super().__init__(parent)
        self.setEditable(editable)
        self.setLayoutDirection(_dir())
        self.setCursor(Qt.PointingHandCursor)

        if editable:
            le = self.lineEdit()
            if le is not None:
                le.setStyleSheet(
                    "QLineEdit { background: transparent; border: none; "
                    "padding: 0px; }"
                )
                self._update_text_margins()
                le.installEventFilter(self)

    def _is_rtl(self):
        return self.layoutDirection() == Qt.RightToLeft

    def _update_text_margins(self):
        le = self.lineEdit()
        if le is None:
            return
        if self._is_rtl():
            le.setTextMargins(self.ARROW_AREA, 0, 8, 0)
        else:
            le.setTextMargins(8, 0, self.ARROW_AREA, 0)

    def setStyleSheet(self, qss):
        qss = (qss or "")
        if not self._is_rtl():
            qss = qss.replace("padding-left: 26px;", "__PL__;")
            qss = qss.replace("padding-right: 10px;", "padding-right: 26px;")
            qss = qss.replace("__PL__;", "padding-left: 10px;")
        super().setStyleSheet(qss + self._ARROW_QSS)

    def eventFilter(self, obj, event):
        le = self.lineEdit()
        if obj is le and event.type() == QEvent.MouseButtonPress:
            if event.button() == Qt.LeftButton:
                try:
                    x = event.position().x()
                except Exception:
                    x = event.pos().x()
                w = le.width() or self.width()
                if self._is_rtl():
                    if x < self.ARROW_AREA:
                        self.showPopup()
                        return True
                else:
                    if x > w - self.ARROW_AREA:
                        self.showPopup()
                        return True
        return super().eventFilter(obj, event)

    def paintEvent(self, e):
        super().paintEvent(e)
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        pm = icon_chevron_down(16).pixmap(QSize(16, 16))
        y = (self.height() - 16) // 2
        x = 4 if self._is_rtl() else self.width() - 16 - 4
        p.drawPixmap(x, y, pm)
        p.end()


# ═════════════════════════════════════════════════════════════════
#  ThemedDialog
# ═════════════════════════════════════════════════════════════════

class ThemedDialog(QDialog):
    def __init__(self, parent, title, message, kind="info",
                 buttons=None):
        if buttons is None:
            buttons = ((tr(T.ok), True, QDialog.Accepted),)
        super().__init__(parent)
        self.setWindowFlags(Qt.Dialog | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setModal(True)
        self.setLayoutDirection(_dir())
        self.setFixedWidth(430)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 16, 16, 16)

        frame = QFrame(); frame.setObjectName("dlgFrame")
        frame.setStyleSheet(f"""
            QFrame#dlgFrame {{
                background-color: #FFFFFF;
                border-radius: 14px;
                border: 1px solid {C_BORDER};
            }}
        """)
        sh = QGraphicsDropShadowEffect()
        sh.setBlurRadius(34); sh.setOffset(0, 8)
        sh.setColor(QColor(0, 0, 0, 85))
        frame.setGraphicsEffect(sh)
        outer.addWidget(frame)

        v = QVBoxLayout(frame); v.setContentsMargins(0, 0, 0, 0); v.setSpacing(0)

        tb = QFrame(); tb.setFixedHeight(42)
        tb.setStyleSheet(f"""
            background-color: {C_DARK_GREEN};
            border-top-left-radius: 14px; border-top-right-radius: 14px;
        """)
        tb_lay = QHBoxLayout(tb)
        tb_lay.setContentsMargins(14, 0, 8, 0); tb_lay.setSpacing(8)
        t = QLabel(title)
        tf = QFont(); tf.setPointSize(10); tf.setBold(True)
        t.setFont(tf); t.setStyleSheet("color: white; background: transparent;")
        tb_lay.addWidget(t); tb_lay.addStretch()
        close_btn = QPushButton()
        close_btn.setIcon(icon_close(14, "#FFFFFF"))
        close_btn.setIconSize(QSize(14, 14))
        close_btn.setFixedSize(28, 24)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.setStyleSheet(f"""
            QPushButton {{
                color: white; background: transparent; border: none;
                font-size: 12px; border-radius: 5px; padding-bottom: 2px;
            }}
            QPushButton:hover {{ background-color: {C_RED}; }}
        """)
        close_btn.clicked.connect(self.reject)
        tb_lay.addWidget(close_btn)
        v.addWidget(tb)

        body = QWidget()
        body_lay = QVBoxLayout(body)
        body_lay.setContentsMargins(24, 22, 24, 18); body_lay.setSpacing(18)

        row = QHBoxLayout(); row.setSpacing(14)
        ic = QLabel()
        ic.setPixmap(icon_status(42, kind).pixmap(QSize(42, 42)))
        ic.setFixedSize(42, 42)
        ic.setStyleSheet("background: transparent;")
        ic.setAlignment(Qt.AlignTop | Qt.AlignHCenter)
        msg = QLabel(message)
        mf = QFont(); mf.setPointSize(10)
        msg.setFont(mf); msg.setWordWrap(True)
        msg.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")
        row.addWidget(ic); row.addWidget(msg, 1)
        body_lay.addLayout(row)

        btn_row = QHBoxLayout(); btn_row.setSpacing(10)
        btn_row.addStretch()
        for text, primary, result in buttons:
            b = QPushButton(text)
            b.setCursor(Qt.PointingHandCursor)
            b.setFixedHeight(38); b.setMinimumWidth(96)
            if primary:
                b.setStyleSheet(f"""
                    QPushButton {{
                        background-color: {C_GREEN};
                        color: white; border: none;
                        border-radius: 8px; font-size: 11px;
                        font-weight: bold; padding: 0 18px;
                    }}
                    QPushButton:hover {{ background-color: {C_DARK_GREEN}; }}
                """)
            else:
                b.setStyleSheet(f"""
                    QPushButton {{
                        background-color: transparent;
                        color: {C_DARK_GREEN};
                        border: 1px solid {C_BORDER};
                        border-radius: 8px; font-size: 11px;
                        padding: 0 18px;
                    }}
                    QPushButton:hover {{
                        background-color: {C_LIGHT_GREEN};
                        border-color: {C_GREEN};
                    }}
                """)
            b.clicked.connect(lambda _=None, r=result: self.done(r))
            btn_row.addWidget(b)
        body_lay.addLayout(btn_row)
        v.addWidget(body)

    @staticmethod
    def info(parent, title, message):
        ThemedDialog(parent, title, message, "info").exec()

    @staticmethod
    def warning(parent, title, message):
        ThemedDialog(parent, title, message, "warning").exec()

    @staticmethod
    def error(parent, title, message):
        ThemedDialog(parent, title, message, "error").exec()

    @staticmethod
    def confirm(parent, title, message):
        d = ThemedDialog(parent, title, message, "question",
                         buttons=((tr(T.yes), True, QDialog.Accepted),
                                  (tr(T.cancel), False, QDialog.Rejected)))
        return d.exec() == QDialog.Accepted

    @staticmethod
    def confirm_delete(parent, title, message):
        d = ThemedDialog(parent, title, message, "delete",
                         buttons=((tr(T.delete), True, QDialog.Accepted),
                                  (tr(T.cancel), False, QDialog.Rejected)))
        return d.exec() == QDialog.Accepted

    def showEvent(self, e):
        super().showEvent(e)
        par = self.parent()
        if par:
            pg = par.window().geometry()
            self.move(pg.center().x() - self.width() // 2,
                      pg.center().y() - self.height() // 2)


# ═════════════════════════════════════════════════════════════════
#  QuickAccessPicker
# ═════════════════════════════════════════════════════════════════

class QuickAccessPicker(QDialog):
    def __init__(self, parent, all_templates, current):
        super().__init__(parent)
        self.setWindowFlags(Qt.Dialog | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setModal(True)
        self.setLayoutDirection(_dir())
        self.setFixedWidth(440)

        self._checks = {}

        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 16, 16, 16)

        frame = QFrame(); frame.setObjectName("dlgFrame")
        frame.setStyleSheet(f"""
            QFrame#dlgFrame {{
                background-color: #FFFFFF;
                border-radius: 14px;
                border: 1px solid {C_BORDER};
            }}
        """)
        sh = QGraphicsDropShadowEffect()
        sh.setBlurRadius(34); sh.setOffset(0, 8)
        sh.setColor(QColor(0, 0, 0, 85))
        frame.setGraphicsEffect(sh)
        outer.addWidget(frame)

        v = QVBoxLayout(frame); v.setContentsMargins(0, 0, 0, 0); v.setSpacing(0)

        tb = QFrame(); tb.setFixedHeight(42)
        tb.setStyleSheet(f"""
            background-color: {C_DARK_GREEN};
            border-top-left-radius: 14px; border-top-right-radius: 14px;
        """)
        tb_lay = QHBoxLayout(tb)
        tb_lay.setContentsMargins(14, 0, 8, 0); tb_lay.setSpacing(8)
        t = QLabel(tr(T.add_to_qa))
        tf = QFont(); tf.setPointSize(10); tf.setBold(True)
        t.setFont(tf); t.setStyleSheet("color: white; background: transparent;")
        tb_lay.addWidget(t); tb_lay.addStretch()
        close_btn = QPushButton()
        close_btn.setIcon(icon_close(14, "#FFFFFF"))
        close_btn.setIconSize(QSize(14, 14))
        close_btn.setFixedSize(28, 24)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.setStyleSheet(f"""
            QPushButton {{
                color: white; background: transparent; border: none;
                font-size: 12px; border-radius: 5px; padding-bottom: 2px;
            }}
            QPushButton:hover {{ background-color: {C_RED}; }}
        """)
        close_btn.clicked.connect(self.reject)
        tb_lay.addWidget(close_btn)
        v.addWidget(tb)

        body = QWidget()
        body_lay = QVBoxLayout(body)
        body_lay.setContentsMargins(20, 16, 20, 16); body_lay.setSpacing(12)

        hint = QLabel(tr(T.select_qa_hint))
        hint.setStyleSheet(f"color: {C_MUTED}; background: transparent; font-size: 10px;")
        hint.setWordWrap(True)
        body_lay.addWidget(hint)

        if not all_templates:
            empty = QLabel(tr(T.no_templates_yet))
            empty.setStyleSheet(
                f"color: {C_MUTED}; background: transparent; "
                f"font-size: 11px; padding: 20px;")
            empty.setAlignment(Qt.AlignCenter)
            body_lay.addWidget(empty)
        else:
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setFrameShape(QFrame.NoFrame)
            scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            scroll.setStyleSheet(f"""
                QScrollArea {{ background: transparent; border: none; }}
                QScrollBar:vertical {{
                    background: transparent; width: 6px;
                }}
                QScrollBar::handle:vertical {{
                    background: {C_BORDER}; border-radius: 3px; min-height: 24px;
                }}
                QScrollBar::handle:vertical:hover {{ background: {C_GREEN}; }}
                QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                    height: 0;
                }}
            """)

            content = QWidget()
            content.setStyleSheet("background: transparent;")
            cl = QVBoxLayout(content)
            cl.setContentsMargins(0, 0, 0, 0); cl.setSpacing(4)

            for tpl in all_templates:
                name = (tpl.get("name") or "").strip()
                if not name:
                    continue
                cb = QCheckBox(name)
                cb.setChecked(name in current)
                cb.setCursor(Qt.PointingHandCursor)
                cb.setStyleSheet(f"""
                    QCheckBox {{
                        color: {C_DARK_GREEN};
                        font-size: 11px;
                        padding: 6px 8px;
                        border-radius: 6px;
                    }}
                    QCheckBox:hover {{
                        background-color: {C_HOVER};
                    }}
                    QCheckBox::indicator {{
                        width: 16px; height: 16px;
                        border: 1.5px solid {C_BORDER};
                        border-radius: 4px;
                        background: #FFFFFF;
                    }}
                    QCheckBox::indicator:hover {{
                        border-color: {C_GREEN};
                    }}
                    QCheckBox::indicator:checked {{
                        background-color: {C_GREEN};
                        border-color: {C_GREEN};
                    }}
                """)
                cl.addWidget(cb)
                self._checks[name] = cb
            cl.addStretch()

            scroll.setWidget(content)
            scroll.setMinimumHeight(200)
            scroll.setMaximumHeight(320)
            body_lay.addWidget(scroll)

        btn_row = QHBoxLayout(); btn_row.setSpacing(10)
        btn_row.addStretch()

        ok_btn = QPushButton(tr(T.save))
        ok_btn.setCursor(Qt.PointingHandCursor)
        ok_btn.setFixedHeight(38); ok_btn.setMinimumWidth(100)
        ok_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {C_GREEN}; color: white;
                border: none; border-radius: 8px;
                font-size: 11px; font-weight: bold;
                padding: 0 18px;
            }}
            QPushButton:hover {{ background-color: {C_DARK_GREEN}; }}
        """)
        ok_btn.clicked.connect(self.accept)
        btn_row.addWidget(ok_btn)

        cancel_btn = QPushButton(tr(T.cancel))
        cancel_btn.setCursor(Qt.PointingHandCursor)
        cancel_btn.setFixedHeight(38); cancel_btn.setMinimumWidth(100)
        cancel_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                color: {C_DARK_GREEN};
                border: 1px solid {C_BORDER};
                border-radius: 8px; font-size: 11px;
                padding: 0 18px;
            }}
            QPushButton:hover {{
                background-color: {C_LIGHT_GREEN};
                border-color: {C_GREEN};
            }}
        """)
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)
        body_lay.addLayout(btn_row)

        v.addWidget(body)

    def selected_names(self):
        return [name for name, cb in self._checks.items() if cb.isChecked()]

    def showEvent(self, e):
        super().showEvent(e)
        par = self.parent()
        if par:
            pg = par.window().geometry()
            self.move(pg.center().x() - self.width() // 2,
                      pg.center().y() - self.height() // 2)


# ═════════════════════════════════════════════════════════════════
#  JSON management
# ═════════════════════════════════════════════════════════════════

def load_templates_raw():
    if not TEMPLATES_JSON.exists():
        return []
    try:
        with open(TEMPLATES_JSON, "r", encoding="utf-8") as f:
            data = load(f)
        return data if isinstance(data, list) else []
    except Exception as e:
        print(f"[error] reading JSON: {e}")
        return []


def load_templates():
    return [t for t in load_templates_raw()
            if isinstance(t, dict) and "name" in t and "address" in t]


def save_templates(templates):
    WORDCLONER_DIR.mkdir(parents=True, exist_ok=True)
    with open(TEMPLATES_JSON, "w", encoding="utf-8") as f:
        dump(templates, f, ensure_ascii=False, indent=4)


# ═════════════════════════════════════════════════════════════════
#  Word processing
# ═════════════════════════════════════════════════════════════════

def normalize_quotes(text):
    return (text
            .replace("\u201c", '"').replace("\u201d", '"')
            .replace("\u201e", '"').replace("\u201f", '"')
            .replace("\u2033", '"'))


def _get_wrapper():
    return SETTINGS.get("placeholder_wrapper", '##')


class WrapperCountError(ValueError):
    pass


def _run_size_pt(run):
    try:
        if run.font.size is not None:
            return run.font.size.pt
    except Exception:
        pass
    return None


def compute_auto_font(runs, char_map, text, ph_start, ph_end):
    def collect_before():
        sizes = []
        i = ph_start - 1
        while i >= 0:
            while i >= 0 and text[i].isspace():
                i -= 1
            if i < 0:
                break
            while i >= 0 and not text[i].isspace():
                i -= 1
            word_start = i + 1
            ri, _ = char_map[word_start]
            s = _run_size_pt(runs[ri])
            if s is not None:
                sizes.append(s)
        return sizes

    def collect_after():
        sizes = []
        i = ph_end
        n = len(text)
        while i < n:
            while i < n and text[i].isspace():
                i += 1
            if i >= n:
                break
            word_start = i
            while i < n and not text[i].isspace():
                i += 1
            ri, _ = char_map[word_start]
            s = _run_size_pt(runs[ri])
            if s is not None:
                sizes.append(s)
        return sizes

    before = collect_before()
    after = collect_after()

    checked = []
    max_d = max(len(before), len(after))
    for d in range(max_d):
        b = before[d] if d < len(before) else None
        a = after[d] if d < len(after) else None
        if b is not None:
            checked.append(b)
        if a is not None:
            checked.append(a)
        if b is not None and a is not None and abs(b - a) < 0.01:
            return b
        if b is None or a is None:
            break

    if not checked:
        return None
    c = Counter(round(x, 1) for x in checked)
    return c.most_common(1)[0][0]


def _nearest_run(runs, char_map, text, ph_start, ph_end):
    i = ph_start - 1
    while i >= 0 and text[i].isspace():
        i -= 1
    if i >= 0:
        while i >= 0 and not text[i].isspace():
            i -= 1
        word_start = i + 1
        return runs[char_map[word_start][0]]
    i = ph_end
    n = len(text)
    while i < n and text[i].isspace():
        i += 1
    if i < n:
        return runs[char_map[i][0]]
    return None


def _copy_run_fonts(src_run, dst_run):
    if not HAS_DOCX:
        return
    try:
        src_rPr = src_run._element.find(qn('w:rPr'))
        if src_rPr is None:
            return
        src_fonts = src_rPr.find(qn('w:rFonts'))
        if src_fonts is None:
            return
        dst_rPr = dst_run._element.get_or_add_rPr()
        existing = dst_rPr.find(qn('w:rFonts'))
        if existing is not None:
            dst_rPr.remove(existing)
        dst_rPr.insert(0, deepcopy(src_fonts))
    except Exception:
        pass


def _template_paragraphs_in_document_order(doc):
    """Return body paragraphs and table-cell paragraphs in reading order.

    Placeholder replacement already processes body paragraphs and top-level
    tables. Keeping extraction in the same document order ensures CLI values
    map predictably even when a template uses table-based layouts.
    """
    paragraphs = []
    seen = set()

    def add_paragraph(paragraph):
        element = paragraph._p
        if element in seen:
            return
        seen.add(element)
        paragraphs.append(paragraph)

    for block in doc.element.body.iterchildren():
        if block.tag == qn("w:p"):
            add_paragraph(DocxParagraph(block, doc))
        elif block.tag == qn("w:tbl"):
            table = DocxTable(block, doc)
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        add_paragraph(paragraph)

    return paragraphs


def extract_placeholders(docx_path):
    doc = Document(docx_path)
    paragraphs = _template_paragraphs_in_document_order(doc)
    para_texts = [p.text for p in paragraphs]

    para_starts = []
    offset = 0
    for t in para_texts:
        para_starts.append(offset)
        offset += len(t) + 1

    raw_text = "\n".join(para_texts)
    text = normalize_quotes(raw_text)

    wrapper = normalize_quotes(_get_wrapper())
    wlen = len(wrapper)
    if wlen == 0:
        return {}

    positions = []
    start = 0
    while True:
        idx = text.find(wrapper, start)
        if idx == -1:
            break
        positions.append(idx)
        start = idx + wlen

    if len(positions) % 2 != 0:
        raise WrapperCountError(tr(T.var_count_error_detail,
                                    w=wrapper, n=len(positions)))

    props = {}
    for i in range(0, len(positions) - 1, 2):
        open_pos = positions[i]
        close_pos = positions[i + 1]
        term_start = open_pos + wlen
        term = text[term_start:close_pos]
        term_key = (term or "").strip()
        if not term_key or term_key in props:
            continue

        pi = 0
        for k in range(len(para_starts) - 1, -1, -1):
            if para_starts[k] <= open_pos:
                pi = k
                break

        auto_size = None
        font_name = None

        if pi < len(para_texts):
            para_start = para_starts[pi]
            para_text_norm = normalize_quotes(para_texts[pi])
            local_start = open_pos - para_start
            local_end = close_pos + wlen - para_start

            if (0 <= local_start < len(para_text_norm)
                    and local_end <= len(para_text_norm)):
                p = paragraphs[pi]
                runs = list(p.runs)
                if runs:
                    parts = []
                    char_map = []
                    for ri, r in enumerate(runs):
                        for ci, ch in enumerate(r.text):
                            parts.append(ch)
                            char_map.append((ri, ci))
                    para_join = normalize_quotes("".join(parts))
                    if len(para_join) == len(para_text_norm):
                        try:
                            auto_size = compute_auto_font(
                                runs, char_map, para_join,
                                local_start, local_end)
                        except Exception:
                            auto_size = None
                        try:
                            ref = _nearest_run(
                                runs, char_map, para_join,
                                local_start, local_end)
                            if ref is not None:
                                try:
                                    if ref.font.name:
                                        font_name = ref.font.name
                                except Exception:
                                    pass
                        except Exception:
                            pass

        props[term_key] = {
            "start": term_start,
            "auto_size": auto_size,
            "font_name": font_name,
        }
    return props


def replace_placeholders_in_docx(src_path, dst_path, replacements, font_overrides):
    doc = Document(src_path)
    wrapper = normalize_quotes(_get_wrapper())
    wlen = len(wrapper)
    if wlen == 0:
        doc.save(dst_path)
        return

    def process_paragraph(paragraph):
        original_runs = list(paragraph.runs)
        if not original_runs:
            return

        parts = []
        char_map = []
        for ri, r in enumerate(original_runs):
            for ci, ch in enumerate(r.text):
                parts.append(ch)
                char_map.append((ri, ci))
        text = "".join(parts)
        norm = normalize_quotes(text)

        ranges = []
        i = 0
        while i < len(norm):
            start = norm.find(wrapper, i)
            if start == -1:
                break
            end = norm.find(wrapper, start + wlen)
            if end == -1:
                break
            key = norm[start + wlen:end].strip()
            if key in replacements:
                ranges.append((start, end + wlen, key))
                i = end + wlen
            else:
                i = start + wlen

        if not ranges:
            return

        auto_sizes = {}
        auto_refs = {}
        for (start, end, key) in ranges:
            override = font_overrides.get(key, (None, False, False, None))
            if override[0] is None:
                auto_sizes[(start, end, key)] = compute_auto_font(
                    original_runs, char_map, norm, start, end)
            else:
                auto_sizes[(start, end, key)] = override[0]
            auto_refs[(start, end, key)] = _nearest_run(
                original_runs, char_map, norm, start, end)

        segments = []

        def emit_orig(lo, hi):
            if lo >= hi:
                return
            seg_lo = lo
            cur_ri = char_map[lo][0]
            for k in range(lo + 1, hi):
                if char_map[k][0] != cur_ri:
                    segments.append((norm[seg_lo:k], original_runs[cur_ri], False))
                    seg_lo = k
                    cur_ri = char_map[k][0]
            segments.append((norm[seg_lo:hi], original_runs[cur_ri], False))

        pos = 0
        for (start, end, key) in ranges:
            emit_orig(pos, start)
            override = font_overrides.get(key, (None, False, False, None))
            size_pt = auto_sizes[(start, end, key)]
            ref_run = auto_refs.get((start, end, key))
            color_hex = override[3] if len(override) > 3 else None
            value = replacements[key]
            segments.append(
                (value,
                 (size_pt, override[1], override[2], color_hex, ref_run),
                 True)
            )
            pos = end

        if pos < len(norm):
            emit_orig(pos, len(norm))

        for r in original_runs:
            r._element.getparent().remove(r._element)

        for seg_text, font_spec, is_repl in segments:
            if not seg_text:
                continue
            new_run = paragraph.add_run(seg_text)
            if is_repl:
                size_pt, bold, italic, color_hex, ref_run = font_spec
                if ref_run is not None:
                    _copy_run_fonts(ref_run, new_run)
                if size_pt is not None:
                    try:
                        new_run.font.size = Pt(float(size_pt))
                    except Exception:
                        pass
                try:
                    new_run.font.bold = bool(bold)
                    new_run.font.italic = bool(italic)
                except Exception:
                    pass
                if color_hex:
                    try:
                        h = color_hex.lstrip("#").upper()
                        new_run.font.color.rgb = RGBColor.from_string(h)
                    except Exception:
                        pass
            else:
                orig = font_spec
                try:
                    new_run.bold = orig.bold
                    new_run.italic = orig.italic
                    new_run.underline = orig.underline
                    if orig.font.size is not None:
                        new_run.font.size = orig.font.size
                    if orig.font.name is not None:
                        new_run.font.name = orig.font.name
                    if (orig.font.color is not None
                            and orig.font.color.rgb is not None):
                        new_run.font.color.rgb = orig.font.color.rgb
                except Exception:
                    pass

    for p in doc.paragraphs:
        process_paragraph(p)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    process_paragraph(p)

    doc.save(dst_path)


class ExcelImportError(ValueError):
    """A workbook cannot be mapped safely to the current template."""


def _excel_is_blank(value):
    return value is None or (isinstance(value, str) and not value.strip())


def _excel_cell_text(value):
    """Convert Excel cell values to stable, human-readable replacement text."""
    if value is None:
        return ""
    if isinstance(value, datetime):
        if not any((value.hour, value.minute, value.second, value.microsecond)):
            return value.date().isoformat()
        return value.isoformat(sep=" ", timespec="seconds")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def read_excel_records(excel_path, variable_names, header_row=False):
    """Read the active worksheet using filename + ordered-variable columns."""
    if not HAS_EXCEL or load_workbook is None:
        raise ExcelImportError(tr(T.batch_excel_library_missing))

    path = Path(excel_path).expanduser()
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        raise ExcelImportError(tr(T.batch_excel_type_error, path=path))
    if not path.is_file():
        raise ExcelImportError(tr(T.batch_excel_not_found, path=path))

    try:
        workbook = load_workbook(str(path), read_only=True, data_only=True)
    except Exception as exc:
        raise ExcelImportError(tr(T.batch_excel_read_error, err=exc)) from exc

    records = []
    try:
        worksheet = workbook.active
        first_data_row = 2 if header_row else 1
        expected_columns = len(variable_names) + 1
        max_column = max(int(worksheet.max_column or 1), expected_columns)

        for row_number, row in enumerate(
                worksheet.iter_rows(
                    min_row=first_data_row,
                    max_row=int(worksheet.max_row or 0),
                    max_col=max_column,
                    values_only=True,
                ),
                start=first_data_row):
            values = list(row)
            if all(_excel_is_blank(value) for value in values):
                continue

            for offset, extra_value in enumerate(values[expected_columns:]):
                if not _excel_is_blank(extra_value):
                    extra_column = expected_columns + offset + 1
                    raise ExcelImportError(tr(
                        T.batch_extra_columns,
                        row=row_number,
                        column=extra_column,
                        variables=len(variable_names),
                    ))

            output_name = _excel_cell_text(values[0]).strip()
            if not output_name:
                raise ExcelImportError(tr(T.batch_filename_empty, row=row_number))

            replacements = {
                variable: _excel_cell_text(
                    values[index + 1] if index + 1 < len(values) else None
                )
                for index, variable in enumerate(variable_names)
            }
            records.append((row_number, output_name, replacements))
    except ExcelImportError:
        raise
    except Exception as exc:
        raise ExcelImportError(tr(T.batch_excel_read_error, err=exc)) from exc
    finally:
        try:
            workbook.close()
        except Exception:
            pass

    if not records:
        raise ExcelImportError(tr(T.batch_no_data_rows))
    return records


def _save_docx_atomically(src_path, dst_path, replacements, font_overrides,
                          overwrite=False):
    """Write a complete DOCX to a sibling temporary file before committing."""
    dst_path = Path(dst_path)
    dst_path.parent.mkdir(parents=True, exist_ok=True)
    if dst_path.exists() and not overwrite:
        raise FileExistsError(str(dst_path))

    fd, temp_name = tempfile.mkstemp(
        prefix=f".{dst_path.stem}.", suffix=".docx", dir=str(dst_path.parent)
    )
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        replace_placeholders_in_docx(
            str(src_path), str(temp_path), replacements, font_overrides
        )
        if dst_path.exists() and not overwrite:
            raise FileExistsError(str(dst_path))
        os.replace(str(temp_path), str(dst_path))
    finally:
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            pass
    return dst_path


def _safe_batch_filename(raw_name, output_dir, reserved_names):
    name = str(raw_name).strip()
    invalid = '<>:"/\\|?*'
    name = "".join("_" if ch in invalid or ord(ch) < 32 else ch for ch in name)
    name = name.rstrip(" .")
    if not name:
        raise ValueError(tr(T.batch_filename_empty, row="?"))
    if not name.lower().endswith(".docx"):
        name += ".docx"

    candidate = Path(name)
    stem = candidate.stem.rstrip(" .") or "Document"
    suffix = ".docx"
    if stem.upper() in {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}:
        stem = f"_{stem}"

    output_dir = Path(output_dir)
    number = 1
    while True:
        file_name = f"{stem}{suffix}" if number == 1 else f"{stem}_{number}{suffix}"
        identity = file_name.casefold()
        if identity not in reserved_names and not (output_dir / file_name).exists():
            reserved_names.add(identity)
            return file_name
        number += 1


def _font_size_value(size_combo):
    size_text = size_combo.currentText().strip().lower()
    if size_text in ("", "auto", tr(T.auto).strip().lower()):
        return None
    digit_map = str.maketrans(
        "۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩",
        "01234567890123456789",
    )
    normalized = size_text.translate(digit_map).replace(",", ".").replace("٫", ".")
    try:
        return float(normalized)
    except ValueError:
        return None


# ═════════════════════════════════════════════════════════════════
#  TitleBar — buttons ALWAYS on the right, both languages
# ═════════════════════════════════════════════════════════════════

class TitleBar(QFrame):
    def __init__(self, window):
        super().__init__()
        self._win = window
        self._drag = None
        self.setFixedHeight(46)
        self.setObjectName("titleBar")

        self.setLayoutDirection(Qt.LeftToRight)
        
        lay = QHBoxLayout(self)
        lay.setContentsMargins(16, 0, 8, 0)
        lay.setSpacing(10)

        ico = QLabel()
        ico.setPixmap(icon_document(20, "#FFFFFF").pixmap(QSize(20, 20)))
        ico.setStyleSheet("background: transparent;")

        t = QLabel(tr(T.app_title))
        f = QFont(); f.setPointSize(11); f.setBold(True)
        t.setFont(f)
        t.setStyleSheet("color: white; background: transparent;")

        lay.addWidget(ico)
        lay.addWidget(t)
        lay.addStretch()

        b_min = self._btn("", C_GREEN)
        b_max = self._btn("", C_GREEN)
        b_cls = self._btn("", C_RED)
        b_min.setIcon(icon_minimize(16, "#FFFFFF"))
        b_max.setIcon(icon_maximize(16, "#FFFFFF"))
        b_cls.setIcon(icon_close(14, "#FFFFFF"))
        b_min.setIconSize(QSize(16, 16))
        b_max.setIconSize(QSize(16, 16))
        b_cls.setIconSize(QSize(14, 14))
        b_min.clicked.connect(self._win.showMinimized)
        b_max.clicked.connect(self._toggle_max)
        b_cls.clicked.connect(self._win.close)

        lay.addWidget(b_min)
        lay.addWidget(b_max)
        lay.addWidget(b_cls)

        self._apply_style()

    def _apply_style(self):
        r = 0 if self._win.isMaximized() else 14
        self.setStyleSheet(f"""
            QFrame#titleBar {{
                background-color: {C_DARK_GREEN};
                border-top-left-radius: {r}px;
                border-top-right-radius: {r}px;
            }}
        """)

    def refresh_radius(self):
        self._apply_style()

    def _btn(self, txt, hover):
        b = QPushButton(txt)
        b.setFixedSize(34, 28)
        b.setCursor(Qt.PointingHandCursor)
        b.setStyleSheet(f"""
            QPushButton {{
                color: white; background: transparent; border: none;
                font-size: 14px; border-radius: 6px; padding-bottom: 2px;
            }}
            QPushButton:hover {{ background-color: {hover}; }}
        """)
        return b

    def _toggle_max(self):
        self._win.showNormal() if self._win.isMaximized() else self._win.showMaximized()

    def mousePressEvent(self, e):
        if e.button() != Qt.LeftButton:
            return
        try:
            wh = self._win.windowHandle()
            if wh is not None:
                if self._win.isMaximized():
                    self._win.showNormal()
                if wh.startSystemMove():
                    e.accept()
                    return
        except Exception:
            pass
        self._drag = e.globalPosition().toPoint() - self._win.frameGeometry().topLeft()

    def mouseMoveEvent(self, e):
        if (e.buttons() & Qt.LeftButton and self._drag
                and not self._win.isMaximized()):
            self._win.move(e.globalPosition().toPoint() - self._drag)

    def mouseReleaseEvent(self, e):
        self._drag = None

    def mouseDoubleClickEvent(self, e):
        self._toggle_max()

# ═════════════════════════════════════════════════════════════════
#  DropdownButton / DragGhost / GripLabel / DragController
# ═════════════════════════════════════════════════════════════════

class DropdownButton(QFrame):
    clicked = Signal()

    def __init__(self):
        super().__init__()
        self.setObjectName("ddbtn")
        self.setFixedHeight(58)
        self.setCursor(Qt.PointingHandCursor)
        self.setStyleSheet(f"""
            QFrame#ddbtn {{
                background-color: #FFFFFF;
                border: 1.5px solid {C_BORDER};
                border-radius: 12px;
            }}
            QFrame#ddbtn:hover {{
                border-color: {C_GREEN};
                background-color: #FBFDFC;
            }}
        """)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(18, 0, 18, 0); lay.setSpacing(12)

        ico = QLabel(); ico.setPixmap(icon_document(22).pixmap(QSize(22, 22)))
        ico.setStyleSheet("background: transparent;")
        ico.setAttribute(Qt.WA_TransparentForMouseEvents)

        self.text = QLabel(tr(T.select_template))
        f = QFont(); f.setPointSize(11)
        self.text.setFont(f)
        self.text.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")
        self.text.setAttribute(Qt.WA_TransparentForMouseEvents)

        chev = QLabel(); chev.setPixmap(icon_chevron_down(16).pixmap(QSize(16, 16)))
        chev.setStyleSheet("background: transparent;")
        chev.setAttribute(Qt.WA_TransparentForMouseEvents)

        lay.addWidget(ico); lay.addWidget(self.text)
        lay.addStretch(); lay.addWidget(chev)

    def set_text(self, s):
        self.text.setText(s)

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self.clicked.emit()


class DragGhost(QWidget):
    def __init__(self, text, icon, width):
        super().__init__(
            None,
            Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint
        )
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setAttribute(Qt.WA_ShowWithoutActivating, True)
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)

        self.setFixedSize(width, 46)

        frame = QFrame(self)
        frame.setGeometry(0, 0, width, 46)
        frame.setStyleSheet(f"""
            QFrame {{
                background-color: #FFFFFF;
                border: 1.5px solid {C_GREEN};
                border-radius: 8px;
            }}
        """)
        lay = QHBoxLayout(frame)
        lay.setContentsMargins(10, 0, 10, 0); lay.setSpacing(8)

        ic = QLabel()
        ic.setPixmap(icon.pixmap(QSize(20, 20)))
        ic.setStyleSheet("background: transparent; border: none;")
        lay.addWidget(ic)

        tx = QLabel(text)
        tx.setStyleSheet(
            f"color: {C_DARK_GREEN}; background: transparent; "
            f"border: none; font-size: 11px;")
        lay.addWidget(tx)
        lay.addStretch()

        self.setWindowOpacity(0.95)


class GripLabel(QLabel):
    def __init__(self, controller, item, parent=None):
        super().__init__(parent)
        self._controller = controller
        self._item = item
        self._dragging = False
        self.setPixmap(icon_grip(16).pixmap(QSize(16, 16)))
        self.setFixedSize(18, 18)
        self.setStyleSheet("background: transparent;")
        self.setCursor(Qt.OpenHandCursor)
        self.setToolTip(tr(T.drag_to_reorder))

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton and self._controller:
            self._dragging = True
            self.setCursor(Qt.ClosedHandCursor)
            self._controller.start(self._item, e.globalPosition().toPoint())

    def mouseMoveEvent(self, e):
        if self._dragging and self._controller:
            self._controller.move(e.globalPosition().toPoint())

    def mouseReleaseEvent(self, e):
        if self._dragging and self._controller:
            self._dragging = False
            self.setCursor(Qt.OpenHandCursor)
            self._controller.end()


class DragController:
    ITEM_H = 48

    def __init__(self, on_reorder=None):
        self.items = []
        self.layout = None
        self.on_reorder = on_reorder
        self.drag_item = None
        self.ghost = None
        self.indicator = None
        self._target_idx = None

    def setup(self, items, layout):
        self.items = items
        self.layout = layout

    def _item_name(self, item):
        try:
            tpl = getattr(item, "_tpl", None) or {}
            return tpl.get("name", "") or getattr(item, "_name", "")
        except Exception:
            return ""

    def start(self, item, global_pos):
        self.drag_item = item
        content = item.parentWidget()
        if content is None:
            return

        name = self._item_name(item)
        self.ghost = DragGhost(name, icon_document(20), item.width())
        self.ghost.move(global_pos.x() - self.ghost.width() // 2,
                        global_pos.y() - 23)
        self.ghost.show()
        self.ghost.raise_()

        self.indicator = QFrame(content)
        self.indicator.setStyleSheet(
            f"background-color: {C_GREEN}; border: none; border-radius: 1px;")
        self.indicator.setFixedHeight(2)
        self.indicator.hide()

        item.setStyleSheet(f"""
            QFrame#it {{
                background-color: {C_HOVER};
                border-radius: 8px;
                border: 1px dashed {C_BORDER};
            }}
            QFrame#qaitem {{
                background-color: {C_HOVER};
                border-radius: 10px;
                border: 1px dashed {C_BORDER};
            }}
        """)

    def move(self, global_pos):
        if not self.drag_item or self.ghost is None:
            return

        self.ghost.move(global_pos.x() - self.ghost.width() // 2,
                        global_pos.y() - 23)

        content = self.drag_item.parentWidget()
        if content is None:
            return
        local = content.mapFromGlobal(global_pos)
        mouse_y = local.y()
        n = len(self.items)
        target_idx = max(0, min(n, int((mouse_y + self.ITEM_H // 2) // self.ITEM_H)))
        self._target_idx = target_idx

        y = max(0, target_idx * self.ITEM_H - 1)
        self.indicator.setGeometry(4, y, max(10, content.width() - 8), 2)
        self.indicator.show()
        self.indicator.raise_()

    def end(self):
        if not self.drag_item:
            return

        target_idx = self._target_idx

        if self.ghost is not None:
            self.ghost.hide()
            self.ghost.deleteLater()
            self.ghost = None
        if self.indicator is not None:
            self.indicator.hide()
            self.indicator.deleteLater()
            self.indicator = None

        self.drag_item.setStyleSheet(f"""
            QFrame#it {{ background-color: transparent; border-radius: 8px; }}
            QFrame#it:hover {{ background-color: {C_HOVER}; }}
            QFrame#qaitem {{
                background-color: #FFFFFF;
                border: 1.5px solid {C_BORDER};
                border-radius: 10px;
            }}
            QFrame#qaitem:hover {{
                border-color: {C_GREEN};
                background-color: {C_HOVER};
            }}
        """)

        if target_idx is not None:
            self._move_item_to(self.drag_item, target_idx)
            if self.on_reorder:
                self.on_reorder(list(self.items))

        self.drag_item = None
        self._target_idx = None

    def _move_item_to(self, item, target_idx):
        items = self.items
        try:
            cur_idx = items.index(item)
        except ValueError:
            return
        if target_idx > cur_idx:
            target_idx -= 1
        if target_idx == cur_idx:
            return
        items.pop(cur_idx)
        items.insert(target_idx, item)
        for it in items:
            self.layout.removeWidget(it)
        for i, it in enumerate(items):
            self.layout.insertWidget(i, it)


class QuickAccessDragController(DragController):
    ITEM_H = 44


class PopupItem(QFrame):
    clicked = Signal()
    delete_clicked = Signal()

    def __init__(self, icon: QIcon, text: str, *, is_add=False,
                 show_delete=False, show_grip=False,
                 drag_controller=None):
        super().__init__()
        self.setObjectName("it")
        self.setFixedHeight(46)
        self.setCursor(Qt.PointingHandCursor)
        self._tpl = None
        self.grip_label = None

        bg       = C_LIGHT_GREEN if is_add else "transparent"
        bg_hover = C_ACCENT1 if is_add else C_HOVER
        self.setStyleSheet(f"""
            QFrame#it {{ background-color: {bg}; border-radius: 8px; }}
            QFrame#it:hover {{ background-color: {bg_hover}; }}
        """)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(8, 0, 10, 0); lay.setSpacing(8)

        if show_grip and drag_controller is not None:
            grip = GripLabel(drag_controller, self)
            self.grip_label = grip
            lay.addWidget(grip)

        ic = QLabel(); ic.setPixmap(icon.pixmap(QSize(26, 26)))
        ic.setStyleSheet("background: transparent;")
        ic.setAttribute(Qt.WA_TransparentForMouseEvents)
        lay.addWidget(ic)

        tx = QLabel(text)
        f = QFont(); f.setPointSize(11)
        if is_add: f.setBold(True)
        tx.setFont(f)
        tx.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")
        tx.setAttribute(Qt.WA_TransparentForMouseEvents)
        lay.addWidget(tx)
        lay.addStretch()

        self.del_btn = None
        if show_delete:
            self.del_btn = QPushButton()
            self.del_btn.setIcon(icon_trash(18))
            self.del_btn.setIconSize(QSize(18, 18))
            self.del_btn.setFixedSize(30, 30)
            self.del_btn.setCursor(Qt.PointingHandCursor)
            self.del_btn.setVisible(False)
            self.del_btn.setStyleSheet("""
                QPushButton { background: transparent; border: none; border-radius: 6px; }
                QPushButton:hover { background-color: #FCE8E6; }
            """)
            self.del_btn.clicked.connect(self.delete_clicked.emit)
            lay.addWidget(self.del_btn)

    def enterEvent(self, e):
        if self.del_btn: self.del_btn.setVisible(True)

    def leaveEvent(self, e):
        if self.del_btn: self.del_btn.setVisible(False)

    def mousePressEvent(self, e):
        if e.button() != Qt.LeftButton: return
        if self.del_btn and self.del_btn.isVisible():
            if self.del_btn.geometry().contains(e.position().toPoint()):
                return
        self.clicked.emit()


class DropdownPopup(QWidget):
    add_clicked = Signal()
    template_selected = Signal(dict)
    template_deleted = Signal(dict)
    reordered = Signal(list)

    ADD_ITEM_H = 46
    SEP_H      = 10
    PADDING_H  = 16
    SEARCH_H   = 44
    TEMPLATE_H = 48

    def __init__(self, templates, parent=None):
        super().__init__(
            parent,
            Qt.Popup | Qt.FramelessWindowHint | Qt.NoDropShadowWindowHint
        )
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setLayoutDirection(_dir())

        self._templates = templates
        self._scroll = None
        self._content_h = 0
        self._items = []
        self._max_scroll_h = 320

        self._drag = DragController(on_reorder=self._on_reorder)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)

        frame = QFrame(); frame.setObjectName("popupFrame")
        frame.setStyleSheet(f"""
            QFrame#popupFrame {{
                background-color: #FFFFFF;
                border: 1px solid {C_BORDER};
                border-radius: 14px;
            }}
        """)
        outer.addWidget(frame)

        inner = QVBoxLayout(frame)
        inner.setContentsMargins(8, 8, 8, 8); inner.setSpacing(2)

        add_item = PopupItem(icon_plus_circle(28), tr(T.add_letter), is_add=True)
        add_item.clicked.connect(self.add_clicked.emit)
        inner.addWidget(add_item)

        search_input = QLineEdit()
        search_input.setPlaceholderText(tr(T.search_template))
        search_input.setFixedHeight(34)
        search_input.setLayoutDirection(_dir())
        search_input.setStyleSheet(f"""
            QLineEdit {{
                background-color: {C_HOVER};
                border: 1px solid {C_BORDER};
                border-radius: 8px;
                padding: 0 10px;
                color: {C_DARK_GREEN};
                font-size: 11px;
                margin: 4px 2px 4px 2px;
            }}
            QLineEdit:focus {{
                border-color: {C_GREEN};
                background-color: #FFFFFF;
            }}
        """)
        search_input.textChanged.connect(self._on_search_changed)
        inner.addWidget(search_input)
        self._search = search_input

        if templates:
            sep = QFrame(); sep.setFixedHeight(1)
            sep.setStyleSheet(f"background-color: {C_BORDER}; margin: 2px 10px;")
            inner.addWidget(sep)

            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setFrameShape(QFrame.NoFrame)
            scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            scroll.setStyleSheet(f"""
                QScrollArea {{ background: transparent; border: none; }}
                QScrollArea > QWidget > QWidget {{ background: transparent; }}
                QScrollBar:vertical {{
                    background: transparent; width: 6px; margin: 2px 0 2px 0;
                    border: none;
                }}
                QScrollBar::handle:vertical {{
                    background: {C_BORDER}; border-radius: 3px; min-height: 24px;
                }}
                QScrollBar::handle:vertical:hover {{ background: {C_GREEN}; }}
                QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                    height: 0; background: transparent;
                }}
                QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
                    background: transparent;
                }}
            """)

            content = QWidget()
            content.setStyleSheet("background: transparent;")
            cl = QVBoxLayout(content)
            cl.setContentsMargins(0, 0, 0, 0); cl.setSpacing(2)

            for tpl in templates:
                it = PopupItem(
                    icon_document(22), tpl.get("name", tr(T.untitled)),
                    show_delete=True, show_grip=True,
                    drag_controller=self._drag)
                it._tpl = tpl
                it.clicked.connect(lambda _=None, t=tpl: self.template_selected.emit(t))
                it.delete_clicked.connect(lambda _=None, t=tpl: self.template_deleted.emit(t))
                cl.addWidget(it)
                self._items.append(it)
            cl.addStretch()

            self._drag.setup(self._items, cl)

            scroll.setWidget(content)
            self._scroll = scroll
            self._content_h = len(templates) * self.TEMPLATE_H + 8
            scroll.setFixedHeight(self._content_h)
            inner.addWidget(scroll)

    def _on_search_changed(self, text):
        text = (text or "").strip().lower()
        is_filtering = bool(text)
        count = 0
        for item in self._items:
            name = ""
            if item._tpl is not None:
                name = (item._tpl.get("name") or "").lower()
            show = (not text) or (text in name)
            item.setVisible(show)
            if show:
                count += 1
            if item.grip_label is not None:
                item.grip_label.setEnabled(not is_filtering)
                item.grip_label.setVisible(not is_filtering)

        if self._scroll is not None:
            new_h = count * self.TEMPLATE_H + 8 if count > 0 else 0
            self._content_h = new_h
            final_h = min(new_h, self._max_scroll_h)
            self._scroll.setFixedHeight(final_h)

    def _on_reorder(self, items):
        new_order = [it._tpl for it in items if it._tpl is not None]
        if not new_order:
            return
        try:
            save_templates(new_order)
        except Exception as e:
            print(f"[error] saving order: {e}")
        self.reordered.emit(new_order)

    def constrain_height(self, available_h):
        if self._scroll is None:
            return
        overhead = (self.ADD_ITEM_H + self.SEP_H + self.PADDING_H
                    + self.SEARCH_H)
        max_scroll = max(60, int(available_h) - overhead)
        self._max_scroll_h = max_scroll
        final_h = min(self._content_h, max_scroll)
        self._scroll.setFixedHeight(final_h)


class TemplateDropdown(QWidget):
    add_requested = Signal()
    template_selected = Signal(dict)
    template_deleted = Signal(dict)

    def __init__(self):
        super().__init__()
        self._templates = []; self._popup = None
        lay = QVBoxLayout(self); lay.setContentsMargins(0, 0, 0, 0)
        self.button = DropdownButton()
        self.button.clicked.connect(self._toggle)
        lay.addWidget(self.button)

    def set_templates(self, tpls):
        self._templates = tpls

    def _toggle(self):
        if self._popup and self._popup.isVisible():
            self._popup.close(); self._popup = None; return

        popup = DropdownPopup(self._templates, self)
        popup.add_clicked.connect(self._on_add)
        popup.template_selected.connect(self._on_select)
        popup.template_deleted.connect(self._on_delete)
        popup.reordered.connect(self._on_reordered)

        popup.setFixedWidth(self.button.width())

        gp = self.button.mapToGlobal(QPoint(0, self.button.height() + 4))

        screen = QGuiApplication.screenAt(gp) or QGuiApplication.primaryScreen()
        avail = screen.availableGeometry()
        available_h = avail.bottom() - gp.y() - 10
        popup.constrain_height(available_h)

        popup.adjustSize()
        popup.move(gp.x(), gp.y())
        popup.show()
        self._popup = popup

    def _close(self):
        if self._popup:
            self._popup.close(); self._popup = None

    def _on_add(self):
        self._close(); self.add_requested.emit()

    def _on_select(self, tpl):
        self._close()
        self.button.set_text(tpl.get("name", tr(T.select_template)))
        self.template_selected.emit(tpl)

    def _on_delete(self, tpl):
        self._close(); self.template_deleted.emit(tpl)

    def _on_reordered(self, new_list):
        self._templates = list(new_list)


# ═════════════════════════════════════════════════════════════════
#  Quick Access
# ═════════════════════════════════════════════════════════════════

class QuickAccessItem(QFrame):
    clicked = Signal()
    removed = Signal()

    def __init__(self, name, drag_controller=None):
        super().__init__()
        self.setObjectName("qaitem")
        self.setFixedHeight(40)
        self.setCursor(Qt.PointingHandCursor)
        self._name = name
        self._tpl = {"name": name}
        self.grip_label = None

        self.setStyleSheet(f"""
            QFrame#qaitem {{
                background-color: #FFFFFF;
                border: 1.5px solid {C_BORDER};
                border-radius: 10px;
            }}
            QFrame#qaitem:hover {{
                border-color: {C_GREEN};
                background-color: {C_HOVER};
            }}
        """)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(4, 0, 8, 0); lay.setSpacing(6)

        if drag_controller is not None:
            grip = GripLabel(drag_controller, self)
            self.grip_label = grip
            lay.addWidget(grip)

        ic = QLabel()
        ic.setPixmap(icon_document(16).pixmap(QSize(16, 16)))
        ic.setStyleSheet("background: transparent;")
        ic.setAttribute(Qt.WA_TransparentForMouseEvents)
        lay.addWidget(ic)

        tx = QLabel(name)
        f = QFont(); f.setPointSize(10)
        tx.setFont(f)
        tx.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")
        tx.setAttribute(Qt.WA_TransparentForMouseEvents)
        lay.addWidget(tx)
        lay.addStretch()

        self.del_btn = QPushButton()
        self.del_btn.setIcon(icon_close(12, C_MUTED))
        self.del_btn.setIconSize(QSize(12, 12))
        self.del_btn.setFixedSize(22, 22)
        self.del_btn.setCursor(Qt.PointingHandCursor)
        self.del_btn.setToolTip(tr(T.remove_from_qa))
        self.del_btn.setVisible(False)
        self.del_btn.setStyleSheet(f"""
            QPushButton {{
                color: {C_MUTED};
                background: transparent;
                border: none;
                font-size: 11px;
                border-radius: 11px;
            }}
            QPushButton:hover {{
                background-color: {C_RED};
                color: white;
            }}
        """)
        self.del_btn.clicked.connect(self.removed.emit)
        lay.addWidget(self.del_btn)

    def enterEvent(self, e):
        self.del_btn.setVisible(True)

    def leaveEvent(self, e):
        self.del_btn.setVisible(False)

    def mousePressEvent(self, e):
        if e.button() != Qt.LeftButton:
            return
        if self.del_btn.isVisible() and \
                self.del_btn.geometry().contains(e.position().toPoint()):
            return
        self.clicked.emit()


class QuickAccessWidget(QWidget):
    template_selected = Signal(dict)
    add_requested = Signal()
    quick_access_changed = Signal()

    def __init__(self):
        super().__init__()
        self.setStyleSheet("background: transparent;")
        self._templates = []
        self._quick_names = []

        self._drag = QuickAccessDragController(on_reorder=self._on_reorder)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0); lay.setSpacing(6)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0); header.setSpacing(8)

        title = QLabel(tr(T.quick_access))
        tf = QFont(); tf.setPointSize(11); tf.setBold(True)
        title.setFont(tf)
        title.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")
        header.addWidget(title)
        header.addStretch()

        self.add_btn = QPushButton()
        self.add_btn.setIcon(icon_plus_white(14))
        self.add_btn.setIconSize(QSize(14, 14))
        self.add_btn.setFixedSize(26, 26)
        self.add_btn.setCursor(Qt.PointingHandCursor)
        self.add_btn.setToolTip(tr(T.add_template_to_qa))
        self.add_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {C_GREEN};
                border: none;
                border-radius: 13px;
                padding: 0px;
            }}
            QPushButton:hover {{ background-color: {C_DARK_GREEN}; }}
        """)
        self.add_btn.clicked.connect(self.add_requested.emit)
        header.addWidget(self.add_btn)
        lay.addLayout(header)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scroll.setStyleSheet(f"""
            QScrollArea {{ background: transparent; border: none; }}
            QScrollArea > QWidget > QWidget {{ background: transparent; }}
            QScrollBar:vertical {{
                background: transparent; width: 6px;
            }}
            QScrollBar::handle:vertical {{
                background: {C_BORDER}; border-radius: 3px; min-height: 24px;
            }}
            QScrollBar::handle:vertical:hover {{ background: {C_GREEN}; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0;
            }}
        """)

        self.items_container = QWidget()
        self.items_container.setStyleSheet("background: transparent;")
        self.items_layout = QVBoxLayout(self.items_container)
        self.items_layout.setContentsMargins(0, 0, 4, 0)
        self.items_layout.setSpacing(4)
        self.items_layout.addStretch()

        self.scroll.setWidget(self.items_container)
        lay.addWidget(self.scroll)

        self.empty_lbl = QLabel(tr(T.no_qa_templates))
        self.empty_lbl.setAlignment(Qt.AlignCenter)
        self.empty_lbl.setStyleSheet(
            f"color: {C_MUTED}; background: transparent; "
            f"font-size: 10px; padding: 10px;")
        lay.addWidget(self.empty_lbl)

        self._rebuild()

    def set_templates(self, tpls):
        self._templates = tpls or []
        self._rebuild()

    def set_quick_names(self, names):
        valid = {t.get("name", "") for t in self._templates}
        self._quick_names = [n for n in (names or []) if n in valid]
        self._rebuild()

    def _rebuild(self):
        while self.items_layout.count() > 1:
            item = self.items_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

        tpl_map = {t.get("name", ""): t for t in self._templates}
        shown = [n for n in self._quick_names if n in tpl_map]

        if not shown:
            self.scroll.setVisible(False)
            self.empty_lbl.setVisible(True)
            self.scroll.setFixedHeight(0)
            self._drag.setup([], None)
            return

        self.empty_lbl.setVisible(False)
        self.scroll.setVisible(True)

        qa_items = []
        for name in shown:
            tpl = tpl_map[name]
            it = QuickAccessItem(name, drag_controller=self._drag)
            it.clicked.connect(lambda _=None, t=tpl: self.template_selected.emit(t))
            it.removed.connect(lambda _=None, n=name: self._remove_name(n))
            self.items_layout.insertWidget(self.items_layout.count() - 1, it)
            qa_items.append(it)

        self._drag.setup(qa_items, self.items_layout)

        total_h = len(shown) * 44 + 4
        self.scroll.setFixedHeight(min(total_h, 176))

    def _remove_name(self, name):
        cur = list(SETTINGS.get("quick_access", []))
        if name in cur:
            cur.remove(name)
        SETTINGS["quick_access"] = cur
        save_settings()
        self._quick_names = [n for n in self._quick_names if n != name]
        self._rebuild()
        self.quick_access_changed.emit() 

    def _on_reorder(self, items):
        new_order = [getattr(it, "_name", "") for it in items
                     if getattr(it, "_name", "")]
        if not new_order:
            return
        SETTINGS["quick_access"] = new_order
        save_settings()
        self._quick_names = new_order
        self.quick_access_changed.emit() 


# ═════════════════════════════════════════════════════════════════
#  DropZone
# ═════════════════════════════════════════════════════════════════

class DropZone(QFrame):
    file_selected = Signal(str)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setObjectName("dropzone")
        self._file_path = None
        self.setMinimumHeight(180)
        self._apply_style(active=False)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 30, 24, 30); lay.setSpacing(14)

        self.icon_lbl = QLabel()
        self.icon_lbl.setPixmap(icon_upload(56).pixmap(QSize(56, 56)))
        self.icon_lbl.setAlignment(Qt.AlignCenter)
        self.icon_lbl.setStyleSheet("background: transparent; border: none;")
        self.icon_lbl.setAttribute(Qt.WA_TransparentForMouseEvents)

        self.text_lbl = QLabel(tr(T.drop_zone_hint))
        self.text_lbl.setAlignment(Qt.AlignCenter)
        f = QFont(); f.setPointSize(11)
        self.text_lbl.setFont(f)
        self.text_lbl.setStyleSheet(f"color: {C_MUTED}; background: transparent; border: none;")
        self.text_lbl.setAttribute(Qt.WA_TransparentForMouseEvents)

        lay.addStretch(); lay.addWidget(self.icon_lbl)
        lay.addWidget(self.text_lbl); lay.addStretch()

    def _apply_style(self, active=False):
        border_color = C_GREEN if active else C_BORDER
        bg = C_LIGHT_GREEN if active else "#FAFCFA"
        self.setStyleSheet(f"""
            QFrame#dropzone {{
                background-color: {bg};
                border: 2px dashed {border_color};
                border-radius: 14px;
            }}
            QFrame#dropzone:hover {{
                border-color: {C_GREEN};
                background-color: {C_HOVER};
            }}
        """)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            urls = e.mimeData().urls()
            if urls and urls[0].toLocalFile().lower().endswith(".docx"):
                e.acceptProposedAction()
                self._apply_style(active=True); return
        e.ignore()

    def dragLeaveEvent(self, e):
        self._apply_style(active=False)

    def dropEvent(self, e):
        self._apply_style(active=False)
        urls = e.mimeData().urls()
        if urls:
            path = urls[0].toLocalFile()
            if path.lower().endswith(".docx"):
                self._set_file(path); e.acceptProposedAction()

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            start_dir = str(self._file_path and Path(self._file_path).parent or DOCUMENTS_DIR)
            path, _ = QFileDialog.getOpenFileName(
                self, tr(T.select_word_file), start_dir, tr(T.word_file_filter))
            if path:
                self._set_file(path)

    def _set_file(self, path):
        self._file_path = path
        self.icon_lbl.setPixmap(icon_check(56).pixmap(QSize(56, 56)))
        self.text_lbl.setText(Path(path).name)
        f = QFont(); f.setPointSize(11); f.setBold(True)
        self.text_lbl.setFont(f)
        self.text_lbl.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent; border: none;")
        self.file_selected.emit(path)

    def file_path(self):
        return self._file_path

    def reset(self):
        self._file_path = None
        self.icon_lbl.setPixmap(icon_upload(56).pixmap(QSize(56, 56)))
        self.text_lbl.setText(tr(T.drop_zone_hint))
        f = QFont(); f.setPointSize(11)
        self.text_lbl.setFont(f)
        self.text_lbl.setStyleSheet(f"color: {C_MUTED}; background: transparent; border: none;")
        self._apply_style(active=False)


class ExcelDropZone(QFrame):
    file_selected = Signal(str)
    SUPPORTED_SUFFIXES = (".xlsx", ".xlsm")

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setObjectName("excelDropZone")
        self._file_path = None
        self.setMinimumHeight(130)
        self._apply_style(active=False)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(8)
        self.icon_lbl = QLabel()
        self.icon_lbl.setPixmap(icon_upload(42).pixmap(QSize(42, 42)))
        self.icon_lbl.setAlignment(Qt.AlignCenter)
        self.icon_lbl.setStyleSheet("background: transparent; border: none;")
        self.icon_lbl.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.text_lbl = QLabel(tr(T.batch_drop_hint))
        self.text_lbl.setAlignment(Qt.AlignCenter)
        self.text_lbl.setWordWrap(True)
        self.text_lbl.setStyleSheet(
            f"color: {C_MUTED}; background: transparent; border: none;"
        )
        self.text_lbl.setAttribute(Qt.WA_TransparentForMouseEvents)
        layout.addStretch()
        layout.addWidget(self.icon_lbl)
        layout.addWidget(self.text_lbl)
        layout.addStretch()

    def _apply_style(self, active=False):
        color = C_GREEN if active else C_BORDER
        background = C_LIGHT_GREEN if active else "#FAFCFA"
        self.setStyleSheet(f"""
            QFrame#excelDropZone {{
                background-color: {background};
                border: 2px dashed {color};
                border-radius: 12px;
            }}
            QFrame#excelDropZone:hover {{
                border-color: {C_GREEN}; background-color: {C_HOVER};
            }}
        """)

    @classmethod
    def _is_supported(cls, path):
        return Path(path).suffix.lower() in cls.SUPPORTED_SUFFIXES

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls and self._is_supported(urls[0].toLocalFile()):
                event.acceptProposedAction()
                self._apply_style(active=True)
                return
        event.ignore()

    def dragLeaveEvent(self, event):
        self._apply_style(active=False)

    def dropEvent(self, event):
        self._apply_style(active=False)
        urls = event.mimeData().urls()
        if urls:
            path = urls[0].toLocalFile()
            if self._is_supported(path):
                self._set_file(path)
                event.acceptProposedAction()

    def mousePressEvent(self, event):
        if event.button() != Qt.LeftButton:
            return
        start_dir = str(
            Path(self._file_path).parent if self._file_path else DOCUMENTS_DIR
        )
        path, _ = QFileDialog.getOpenFileName(
            self,
            tr(T.batch_issue_title),
            start_dir,
            tr(T.batch_excel_filter),
        )
        if path:
            self._set_file(path)

    def _set_file(self, path):
        self._file_path = str(path)
        self.icon_lbl.setPixmap(icon_check(42).pixmap(QSize(42, 42)))
        self.text_lbl.setText(Path(path).name)
        font = QFont(); font.setPointSize(10); font.setBold(True)
        self.text_lbl.setFont(font)
        self.text_lbl.setStyleSheet(
            f"color: {C_DARK_GREEN}; background: transparent; border: none;"
        )
        self.file_selected.emit(self._file_path)

    def file_path(self):
        return self._file_path

    def reset(self):
        self._file_path = None
        self.icon_lbl.setPixmap(icon_upload(42).pixmap(QSize(42, 42)))
        self.text_lbl.setText(tr(T.batch_drop_hint))
        font = QFont(); font.setPointSize(10)
        self.text_lbl.setFont(font)
        self.text_lbl.setStyleSheet(
            f"color: {C_MUTED}; background: transparent; border: none;"
        )
        self._apply_style(active=False)


# ═════════════════════════════════════════════════════════════════
#  Small components
# ═════════════════════════════════════════════════════════════════

class RoundToggle(QPushButton):
    def __init__(self, label, tooltip="", bold=False, italic=False):
        super().__init__(label)
        self.setCheckable(True)
        self.setFixedSize(32, 32)
        self.setCursor(Qt.PointingHandCursor)
        self.setToolTip(tooltip)

        f = QFont()
        f.setPointSize(11)
        f.setBold(bold)
        f.setItalic(italic)
        self.setFont(f)

        self.setStyleSheet(f"""
            QPushButton {{
                background-color: #FFFFFF;
                border: 1.5px solid {C_BORDER};
                border-radius: 16px;
                color: {C_MUTED};
            }}
            QPushButton:hover {{
                border-color: {C_GREEN};
                color: {C_DARK_GREEN};
            }}
            QPushButton:checked {{
                background-color: {C_GREEN};
                border-color: {C_GREEN};
                color: white;
            }}
        """)


FONT_SIZES = ["Auto", "8", "9", "10", "11", "12", "14", "16",
              "18", "20", "22", "24", "26", "28", "36", "48", "72"]


class FontSizeCombo(ComboWithArrow):
    def __init__(self):
        super().__init__(editable=True)
        self.setFixedHeight(32)
        self.setMinimumWidth(100)
        self.addItems([tr(T.auto)] + FONT_SIZES[1:])
        self.setCurrentText(tr(T.auto))
        self.setToolTip(tr(T.font_size_tooltip))

        self.setStyleSheet(f"""
            QComboBox {{
                background-color: #FFFFFF;
                border: 1.5px solid {C_BORDER};
                border-radius: 8px;
                color: {C_DARK_GREEN};
                font-size: 11px;
            }}
            QComboBox:focus {{ border-color: {C_GREEN}; }}
            QComboBox QAbstractItemView {{
                background-color: #FFFFFF;
                border: 1px solid {C_BORDER};
                border-radius: 6px;
                padding: 2px;
                outline: none;
                color: {C_DARK_GREEN};
                selection-background-color: {C_LIGHT_GREEN};
                selection-color: {C_DARK_GREEN};
                max-height: 220px;
            }}
            QComboBox QAbstractItemView::item {{
                padding: 4px 8px; border-radius: 4px; min-height: 20px;
            }}
            QComboBox QAbstractItemView::item:hover {{
                background-color: {C_HOVER};
            }}
        """)

        le = self.lineEdit()
        if le is not None:
            size_pattern = QRegularExpression(
                r"(?:[0-9۰-۹٠-٩]+(?:[.,٫][0-9۰-۹٠-٩]*)?|[.,٫][0-9۰-۹٠-٩]+)?"
            )
            le.setValidator(QRegularExpressionValidator(size_pattern, le))
            le.editingFinished.connect(self._ensure_not_empty)

    def _ensure_not_empty(self):
        le = self.lineEdit()
        if le is None:
            return
        if not le.text().strip():
            le.setText(tr(T.auto))

    def eventFilter(self, obj, event):
        if super().eventFilter(obj, event):
            return True

        le = self.lineEdit()
        if obj is le:
            et = event.type()
            if et == QEvent.MouseButtonPress:
                if le.text().strip():
                    QTimer.singleShot(0, le.clear)
            elif et == QEvent.FocusOut:
                QTimer.singleShot(0, self._ensure_not_empty)
        return False


# ═════════════════════════════════════════════════════════════════
#  HomePage
# ═════════════════════════════════════════════════════════════════

class HomePage(QWidget):
    settings_requested = Signal()
    quick_template_selected = Signal(dict)
    quick_access_changed = Signal()  

    def __init__(self):
        super().__init__()
        self.setStyleSheet("background: transparent;")
        self._templates = []

        outer = QVBoxLayout(self)
        outer.setContentsMargins(40, 18, 40, 24)
        outer.setSpacing(0)

        top_row = QHBoxLayout()
        gear_btn = QPushButton(tr(T.settings_spaced))
        gear_btn.setIcon(icon_gear(20))
        gear_btn.setIconSize(QSize(20, 20))
        gear_btn.setLayoutDirection(_dir())
        gear_btn.setFixedHeight(40)
        gear_btn.setMinimumWidth(120)
        gear_btn.setCursor(Qt.PointingHandCursor)
        gear_btn.setToolTip(tr(T.settings))
        gear_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                color: {C_DARK_GREEN};
                border: 1px solid transparent;
                border-radius: 10px;
                font-size: 11px;
                font-weight: bold;
                padding: 0 14px;
            }}
            QPushButton:hover {{
                background-color: {C_LIGHT_GREEN};
                border-color: {C_BORDER};
            }}
        """)
        gear_btn.clicked.connect(self.settings_requested.emit)

        guide_btn = QPushButton(tr(T.guide))
        guide_btn.setIcon(icon_help(19))
        guide_btn.setIconSize(QSize(19, 19))
        guide_btn.setLayoutDirection(_dir())
        guide_btn.setFixedHeight(38)
        guide_btn.setMinimumWidth(120)
        guide_btn.setCursor(Qt.PointingHandCursor)
        guide_btn.setToolTip(tr(T.guide_tooltip))
        guide_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                color: {C_DARK_GREEN};
                border: 1px solid transparent;
                border-radius: 10px;
                font-size: 11px;
                font-weight: bold;
                padding: 0 14px;
            }}
            QPushButton:hover {{
                background-color: {C_LIGHT_GREEN};
                border-color: {C_BORDER};
            }}
        """)
        guide_btn.clicked.connect(self._open_guide)

        navigation = QVBoxLayout()
        navigation.setContentsMargins(0, 0, 0, 0)
        navigation.setSpacing(2)
        navigation.addWidget(gear_btn)
        navigation.addWidget(guide_btn)
        top_row.addLayout(navigation)
        top_row.addStretch()
        outer.addLayout(top_row)

        outer.addSpacing(4)

        main_scroll = QScrollArea()
        main_scroll.setWidgetResizable(True)
        main_scroll.setFrameShape(QFrame.NoFrame)
        main_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        main_scroll.setStyleSheet(f"""
            QScrollArea {{ background: transparent; border: none; }}
            QScrollBar:vertical {{
                background: transparent; width: 6px;
            }}
            QScrollBar::handle:vertical {{
                background: {C_BORDER}; border-radius: 3px; min-height: 30px;
            }}
            QScrollBar::handle:vertical:hover {{ background: {C_GREEN}; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0;
            }}
        """)

        main_content = QWidget()
        main_content.setStyleSheet("background: transparent;")
        mc_lay = QVBoxLayout(main_content)
        mc_lay.setContentsMargins(0, 0, 0, 0)
        mc_lay.setSpacing(18)

        card = QWidget(); card.setStyleSheet("background: transparent;")
        card.setFixedWidth(440)
        cl = QVBoxLayout(card)
        cl.setContentsMargins(0, 0, 0, 0); cl.setSpacing(8)

        big = QLabel()
        big.setPixmap(icon_document_big(72).pixmap(QSize(72, 72)))
        big.setAlignment(Qt.AlignCenter)
        big.setStyleSheet("background: transparent;")

        title = QLabel(tr(T.issue_letter))
        tf = QFont(); tf.setPointSize(22); tf.setBold(True)
        title.setFont(tf); title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")

        self.dropdown = TemplateDropdown()

        cl.addWidget(big, 0, Qt.AlignHCenter)
        cl.addWidget(title); cl.addSpacing(12); cl.addWidget(self.dropdown)

        center = QHBoxLayout()
        center.addStretch(); center.addWidget(card); center.addStretch()
        mc_lay.addLayout(center)

        qa_wrap = QWidget()
        qa_wrap.setStyleSheet("background: transparent;")
        qa_wrap.setFixedWidth(440)
        qa_lay = QVBoxLayout(qa_wrap)
        qa_lay.setContentsMargins(0, 0, 0, 0)

        self.quick_access = QuickAccessWidget()
        self.quick_access.template_selected.connect(self.quick_template_selected.emit)
        self.quick_access.add_requested.connect(self._on_qa_add)
        qa_lay.addWidget(self.quick_access)
        self.quick_access.quick_access_changed.connect(self.quick_access_changed.emit)
        center_qa = QHBoxLayout()
        center_qa.addStretch(); center_qa.addWidget(qa_wrap); center_qa.addStretch()
        mc_lay.addLayout(center_qa)

        mc_lay.addStretch(1)

        main_scroll.setWidget(main_content)
        outer.addWidget(main_scroll, 1)

    def set_templates(self, tpls):
        self._templates = tpls or []
        self.dropdown.set_templates(self._templates)
        self.quick_access.set_templates(self._templates)
        self.quick_access.set_quick_names(SETTINGS.get("quick_access", []))

    def _on_qa_add(self):
        valid_names = {t.get("name", "") for t in self._templates}
        current = [n for n in SETTINGS.get("quick_access", [])
                   if n in valid_names]
        dlg = QuickAccessPicker(self, self._templates, current)
        if dlg.exec() == QDialog.Accepted:
            new_names = dlg.selected_names()
            SETTINGS["quick_access"] = new_names
            save_settings()
            self.quick_access.set_quick_names(new_names)
            self.quick_access_changed.emit()      

    def _open_guide(self):
        language = SETTINGS.get("language", "fa")
        guide_filename = (
            "English_guide.mp4" if language == "en" else "Persian_guide.mp4"
        )
        program_dir = (
            Path(sys.executable).resolve().parent
            if getattr(sys, "frozen", False)
            else Path(__file__).resolve().parent
        )
        guide_path = program_dir / "Guide" / guide_filename

        if not guide_path.is_file():
            ThemedDialog.warning(
                self, tr(T.error), tr(T.guide_not_found, path=guide_path)
            )
            return

        try:
            if sys.platform == "win32":
                os.startfile(str(guide_path))
            else:
                if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(guide_path))):
                    raise OSError("The operating system did not open the video file.")
        except Exception as exc:
            ThemedDialog.error(
                self, tr(T.error), tr(T.guide_open_error, err=exc)
            )


# ═════════════════════════════════════════════════════════════════
#  AddLetterPage
# ═════════════════════════════════════════════════════════════════

class AddLetterPage(QWidget):
    back_requested = Signal()
    saved = Signal()

    def __init__(self):
        super().__init__()
        self.setStyleSheet("background: transparent;")

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0); root.setSpacing(0)

        header_wrap = QWidget()
        header_wrap.setStyleSheet("background: transparent;")
        hw_lay = QHBoxLayout(header_wrap)
        hw_lay.setContentsMargins(40, 22, 40, 14); hw_lay.setSpacing(10)

        title = QLabel(tr(T.add_new_letter))
        tf = QFont(); tf.setPointSize(15); tf.setBold(True)
        title.setFont(tf)
        title.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")

        back_btn = QPushButton(tr(T.back_spaced))
        back_btn.setIcon(icon_arrow_left(16)); back_btn.setIconSize(QSize(16, 16))
        back_btn.setCursor(Qt.PointingHandCursor); back_btn.setFixedHeight(34)
        back_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent; border: 1px solid {C_BORDER};
                border-radius: 8px; color: {C_DARK_GREEN};
                padding: 0 14px; font-size: 11px;
            }}
            QPushButton:hover {{
                background-color: {C_LIGHT_GREEN};
                border-color: {C_GREEN};
            }}
        """)
        back_btn.clicked.connect(self.back_requested.emit)

        hw_lay.addWidget(title); hw_lay.addStretch(); hw_lay.addWidget(back_btn)
        root.addWidget(header_wrap)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet(f"""
            QScrollArea {{ background: transparent; border: none; }}
            QScrollBar:vertical {{
                background: transparent; width: 8px;
            }}
            QScrollBar::handle:vertical {{
                background: {C_BORDER}; border-radius: 4px; min-height: 30px;
            }}
            QScrollBar::handle:vertical:hover {{ background: {C_GREEN}; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0;
            }}
        """)

        body = QWidget()
        body.setStyleSheet("background: transparent;")
        bl = QVBoxLayout(body)
        bl.setContentsMargins(40, 0, 40, 20)
        bl.setSpacing(14)

        nf = QFont(); nf.setPointSize(10); nf.setBold(True)

        name_lbl = QLabel(tr(T.letter_name))
        name_lbl.setFont(nf)
        name_lbl.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")
        bl.addWidget(name_lbl)

        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText(tr(T.ph_letter_name))
        self.name_input.setFixedHeight(44)
        self.name_input.setStyleSheet(f"""
            QLineEdit {{
                background-color: #FFFFFF;
                border: 1.5px solid {C_BORDER};
                border-radius: 10px;
                padding: 0 14px;
                color: {C_DARK_GREEN};
                font-size: 11px;
            }}
            QLineEdit:focus {{ border-color: {C_GREEN}; }}
        """)
        bl.addWidget(self.name_input)

        drop_lbl = QLabel(tr(T.word_template_file))
        drop_lbl.setFont(nf)
        drop_lbl.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")
        bl.addWidget(drop_lbl)

        self.drop_zone = DropZone()
        bl.addWidget(self.drop_zone)

        bl.addStretch(1)

        scroll.setWidget(body)
        root.addWidget(scroll, 1)

        footer_wrap = QWidget()
        footer_wrap.setStyleSheet("background: transparent;")
        fw_lay = QVBoxLayout(footer_wrap)
        fw_lay.setContentsMargins(40, 8, 40, 22)

        self.confirm_btn = QPushButton(tr(T.confirm_save))
        self.confirm_btn.setCursor(Qt.PointingHandCursor)
        self.confirm_btn.setFixedHeight(48)
        self.confirm_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {C_GREEN}; color: white;
                border: none; border-radius: 10px;
                font-size: 12px; font-weight: bold;
            }}
            QPushButton:hover {{ background-color: {C_DARK_GREEN}; }}
        """)
        self.confirm_btn.clicked.connect(self._on_confirm)
        fw_lay.addWidget(self.confirm_btn)
        root.addWidget(footer_wrap)

    def reset(self):
        self.name_input.clear(); self.drop_zone.reset()

    def _on_confirm(self):
        if not HAS_DOCX:
            ThemedDialog.error(self, tr(T.missing_library), tr(T.no_docx))
            return

        name = self.name_input.text().strip()
        file_path = self.drop_zone.file_path()

        if not name:
            ThemedDialog.warning(self, tr(T.error), tr(T.err_empty_name))
            self.name_input.setFocus(); return
        if not file_path:
            ThemedDialog.warning(self, tr(T.error), tr(T.err_empty_file))
            return

        existing = load_templates_raw()
        if any(t.get("name") == name for t in existing):
            ThemedDialog.warning(self, tr(T.error),
                                 tr(T.err_name_exists, name=name))
            return

        dst = None
        try:
            WORDS_DIR.mkdir(parents=True, exist_ok=True)
            src = Path(file_path)
            dst = WORDS_DIR / src.name
            counter = 1
            while dst.exists():
                dst = WORDS_DIR / f"{src.stem}_{counter}{src.suffix}"
                counter += 1
            copy2(str(src), str(dst))

            try:
                props = extract_placeholders(str(dst))
            except WrapperCountError as we:
                try:
                    dst.unlink()
                except Exception:
                    pass
                ThemedDialog.error(self, tr(T.err_wrapper_count), str(we))
                return

            if not props:
                try:
                    dst.unlink()
                except Exception:
                    pass
                ThemedDialog.warning(
                    self, tr(T.error),
                    tr(T.err_no_vars, w=_get_wrapper()))
                return

            new_obj = {"name": name, "address": str(dst), "Properties": props}
            existing.append(new_obj)
            save_templates(existing)

            ThemedDialog.info(self, tr(T.saved),tr(T.saved_with_n, name=name, n=len(props)))

            self.reset(); self.saved.emit()
        except Exception as ex:
            if dst is not None and dst.exists():
                try:
                    dst.unlink()
                except Exception:
                    pass
            ThemedDialog.error(self, tr(T.error),tr(T.err_generic, err=ex))
            print("[error]", ex)


# ═════════════════════════════════════════════════════════════════
#  EditLetterPage — header static, everything else scrolls
# ═════════════════════════════════════════════════════════════════

class EditLetterPage(QWidget):
    back_requested = Signal()
    batch_requested = Signal()

    def __init__(self):
        super().__init__()
        self.setStyleSheet("background: transparent;")
        self._template = None
        self._fields = {}

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0); root.setSpacing(0)

        header_wrap = QWidget()
        header_wrap.setStyleSheet("background: transparent;")
        hw_lay = QHBoxLayout(header_wrap)
        hw_lay.setContentsMargins(40, 22, 40, 14); hw_lay.setSpacing(10)

        self.title_lbl = QLabel(tr(T.issue_letter))
        tf = QFont(); tf.setPointSize(15); tf.setBold(True)
        self.title_lbl.setFont(tf)
        self.title_lbl.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")

        back_btn = QPushButton(tr(T.back_spaced))
        back_btn.setIcon(icon_arrow_left(16)); back_btn.setIconSize(QSize(16, 16))
        back_btn.setCursor(Qt.PointingHandCursor); back_btn.setFixedHeight(34)
        back_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent; border: 1px solid {C_BORDER};
                border-radius: 8px; color: {C_DARK_GREEN};
                padding: 0 14px; font-size: 11px;
            }}
            QPushButton:hover {{
                background-color: {C_LIGHT_GREEN};
                border-color: {C_GREEN};
            }}
        """)
        back_btn.clicked.connect(self.back_requested.emit)

        self.batch_btn = QPushButton(tr(T.batch_excel_issue))
        self.batch_btn.setIcon(icon_document(17))
        self.batch_btn.setIconSize(QSize(17, 17))
        self.batch_btn.setCursor(Qt.PointingHandCursor)
        self.batch_btn.setFixedHeight(34)
        self.batch_btn.setToolTip(tr(T.batch_excel_issue))
        self.batch_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {C_LIGHT_GREEN};
                border: 1px solid {C_BORDER};
                border-radius: 8px;
                color: {C_DARK_GREEN};
                font-size: 10px; font-weight: bold;
                padding: 0 12px;
            }}
            QPushButton:hover {{
                background-color: {C_HOVER}; border-color: {C_GREEN};
            }}
        """)
        self.batch_btn.clicked.connect(self.batch_requested.emit)

        hw_lay.addWidget(self.title_lbl)
        hw_lay.addStretch()
        hw_lay.addWidget(self.batch_btn)
        hw_lay.addWidget(back_btn)
        root.addWidget(header_wrap)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        self.scroll.setStyleSheet(f"""
            QScrollArea {{ background: transparent; border: none; }}
            QScrollBar:vertical {{
                background: transparent; width: 8px;
            }}
            QScrollBar::handle:vertical {{
                background: {C_BORDER}; border-radius: 4px; min-height: 30px;
            }}
            QScrollBar::handle:vertical:hover {{ background: {C_GREEN}; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0;
            }}
        """)

        self.scroll_body = QWidget()
        self.scroll_body.setStyleSheet("background: transparent;")
        sb = QVBoxLayout(self.scroll_body)
        sb.setContentsMargins(40, 0, 40, 20)
        sb.setSpacing(14)

        nf = QFont(); nf.setPointSize(10); nf.setBold(True)

        fn_lbl = QLabel(tr(T.output_filename))
        fn_lbl.setFont(nf)
        fn_lbl.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")
        sb.addWidget(fn_lbl)

        self.filename_input = QLineEdit()
        self.filename_input.setPlaceholderText(tr(T.ph_filename))
        self.filename_input.setFixedHeight(44)
        self.filename_input.setStyleSheet(f"""
            QLineEdit {{
                background-color: #FFFFFF;
                border: 1.5px solid {C_BORDER};
                border-radius: 10px;
                padding: 0 14px;
                color: {C_DARK_GREEN};
                font-size: 11px;
            }}
            QLineEdit:focus {{ border-color: {C_GREEN}; }}
        """)
        sb.addWidget(self.filename_input)

        vars_lbl = QLabel(tr(T.variable_values))
        vars_lbl.setFont(nf)
        vars_lbl.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")
        sb.addWidget(vars_lbl)

        self.fields_container = QWidget()
        self.fields_container.setStyleSheet("background: transparent;")
        self.fields_layout = QVBoxLayout(self.fields_container)
        self.fields_layout.setContentsMargins(0, 0, 0, 0)
        self.fields_layout.setSpacing(12)
        self.fields_layout.addStretch()

        sb.addWidget(self.fields_container)
        sb.addStretch(1)

        self.scroll.setWidget(self.scroll_body)
        root.addWidget(self.scroll, 1)

        footer_wrap = QWidget()
        footer_wrap.setStyleSheet("background: transparent;")
        fw_lay = QHBoxLayout(footer_wrap)
        fw_lay.setContentsMargins(40, 8, 40, 22)

        self.confirm_btn = QPushButton(tr(T.confirm_issue))
        self.confirm_btn.setCursor(Qt.PointingHandCursor)
        self.confirm_btn.setFixedHeight(46)
        self.confirm_btn.setMinimumWidth(200)
        self.confirm_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {C_GREEN}; color: white;
                border: none; border-radius: 10px;
                font-size: 12px; font-weight: bold;
                padding: 0 28px;
            }}
            QPushButton:hover {{ background-color: {C_DARK_GREEN}; }}
        """)
        self.confirm_btn.clicked.connect(self._on_confirm)
        fw_lay.addWidget(self.confirm_btn)
        fw_lay.addStretch()
        root.addWidget(footer_wrap)

    def load_template(self, tpl):
        self._template = tpl
        self.title_lbl.setText(tr(T.issue_named, name=tpl.get('name', '')))
        self.filename_input.setText(tpl.get("name", ""))

        self._fields.clear()
        while self.fields_layout.count() > 1:
            item = self.fields_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

        props = tpl.get("Properties", {})

        normalized = {}
        for k, v in props.items():
            k2 = (k or "").strip()
            if not k2:
                continue
            if isinstance(v, dict):
                try:
                    start = int(v.get("start", 0))
                except Exception:
                    start = 0
                normalized[k2] = {
                    "start": start,
                    "auto_size": v.get("auto_size"),
                    "font_name": v.get("font_name"),
                }
            else:
                try:
                    start = int(v) if v is not None else 0
                except Exception:
                    start = 0
                normalized[k2] = {
                    "start": start,
                    "auto_size": None,
                    "font_name": None,
                }

        if not normalized:
            empty = QLabel(tr(T.no_variables))
            empty.setStyleSheet(f"color: {C_MUTED}; background: transparent; padding: 8px;")
            empty.setFont(QFont("", 10))
            self.fields_layout.insertWidget(0, empty)
            return

        sorted_items = sorted(normalized.items(),
                              key=lambda kv: kv[1]["start"])
        for idx, (key, info) in enumerate(sorted_items):
            box = self._build_field(
                key,
                show_labels=(idx == 0),
                default_size=None,
            )
            self.fields_layout.insertWidget(self.fields_layout.count() - 1, box)

    def _pick_color(self, key):
        widgets = self._fields.get(key)
        if not widgets:
            return
        btn = widgets.get("color")
        if btn is None:
            return

        cur = btn.property("color") or "#000000"
        dlg = QColorDialog(QColor(cur), self)
        dlg.setOption(QColorDialog.ColorDialogOption.DontUseNativeDialog, True)
        dlg.setOption(QColorDialog.ColorDialogOption.ShowAlphaChannel, False)
        dlg.setLayoutDirection(Qt.LeftToRight)
        dlg.setWindowTitle(tr(T.pick_color))
        dlg.setAutoFillBackground(True)

        palette = dlg.palette()
        palette.setColor(QPalette.ColorRole.Window, QColor("#FFFFFF"))
        palette.setColor(QPalette.ColorRole.WindowText, QColor(C_DARK_GREEN))
        palette.setColor(QPalette.ColorRole.Base, QColor("#FFFFFF"))
        palette.setColor(QPalette.ColorRole.AlternateBase, QColor(C_BG))
        palette.setColor(QPalette.ColorRole.Text, QColor(C_DARK_GREEN))
        palette.setColor(QPalette.ColorRole.Button, QColor(C_LIGHT_GREEN))
        palette.setColor(QPalette.ColorRole.ButtonText, QColor(C_DARK_GREEN))
        palette.setColor(QPalette.ColorRole.Highlight, QColor(C_GREEN))
        palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#FFFFFF"))
        dlg.setPalette(palette)

        dlg.setStyleSheet(f"""
            QColorDialog {{
                background-color: #FFFFFF;
                color: {C_DARK_GREEN};
            }}
            QColorDialog QLabel {{
                color: {C_DARK_GREEN};
                background-color: transparent;
            }}
            QColorDialog QLineEdit,
            QColorDialog QSpinBox,
            QColorDialog QDoubleSpinBox,
            QColorDialog QComboBox {{
                color: {C_DARK_GREEN};
                background-color: #FFFFFF;
                border: 1px solid {C_BORDER};
                border-radius: 4px;
                padding: 3px 6px;
            }}
            QColorDialog QPushButton {{
                color: {C_DARK_GREEN};
                background-color: {C_LIGHT_GREEN};
                border: 1px solid {C_BORDER};
                border-radius: 6px;
                padding: 5px 12px;
            }}
            QColorDialog QPushButton:hover {{
                background-color: {C_HOVER};
                border-color: {C_GREEN};
            }}
        """)

        def localize_color_dialog_labels():
            dlg.setLayoutDirection(Qt.LeftToRight)
            for control in dlg.findChildren(QWidget):
                control.setLayoutDirection(Qt.LeftToRight)
            for label in dlg.findChildren(QLabel):
                label_text = label.text().replace("&", "").strip().rstrip(":").strip().casefold()
                if label_text == "html":
                    label.setText("HEX:")
                    break
            for button in dlg.findChildren(QPushButton):
                button_text = button.text().replace("&", "").strip().casefold()
                if button_text == "ok":
                    button.setText(tr(T.ok))
                elif button_text == "cancel":
                    button.setText(tr(T.cancel))

        QTimer.singleShot(0, localize_color_dialog_labels)

        if dlg.exec():
            col = dlg.currentColor()
            if col.isValid():
                hexc = col.name()
                btn.setProperty("color", hexc)
                apply_color_button_style(btn, hexc)

    def _build_field(self, key, show_labels=False, default_size=None):
        box = QWidget()
        box.setStyleSheet("background: transparent;")
        bl = QVBoxLayout(box)
        bl.setContentsMargins(0, 0, 0, 0); bl.setSpacing(6)

        top = QHBoxLayout()
        top.setSpacing(8)

        lbl = QLabel(key.strip())
        lf = QFont(); lf.setPointSize(10); lf.setBold(True)
        lbl.setFont(lf)
        lbl.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")
        top.addWidget(lbl)
        top.addStretch()

        if show_labels:
            size_lbl = QLabel(tr(T.font_size))
            size_lbl.setStyleSheet(
                f"color: {C_MUTED}; background: transparent; font-size: 10px;")
            top.addWidget(size_lbl)

        size_combo = FontSizeCombo()
        if default_size is not None:
            try:
                if isinstance(default_size, (int, float)):
                    size_combo.setCurrentText(f"{default_size:g}")
                else:
                    size_combo.setCurrentText(str(default_size))
            except Exception:
                pass
        top.addWidget(size_combo)

        bold_wrap = QWidget()
        bold_wrap.setStyleSheet("background: transparent;")
        bold_lay = QHBoxLayout(bold_wrap)
        bold_lay.setContentsMargins(0, 0, 0, 0); bold_lay.setSpacing(6)
        bold_btn = RoundToggle("B", tr(T.bold), bold=True)
        if show_labels:
            bold_lbl = QLabel(tr(T.bold))
            bold_lbl.setStyleSheet(
                f"color: {C_MUTED}; background: transparent; font-size: 10px;")
            bold_lbl.setCursor(Qt.PointingHandCursor)
            bold_lbl.mousePressEvent = lambda _e, b=bold_btn: b.toggle()
            bold_lay.addWidget(bold_lbl)
        bold_lay.addWidget(bold_btn)
        top.addWidget(bold_wrap)

        ital_wrap = QWidget()
        ital_wrap.setStyleSheet("background: transparent;")
        ital_lay = QHBoxLayout(ital_wrap)
        ital_lay.setContentsMargins(0, 0, 0, 0); ital_lay.setSpacing(6)
        italic_btn = RoundToggle("I", tr(T.italic), italic=True)
        if show_labels:
            italic_lbl = QLabel(tr(T.italic))
            italic_lbl.setStyleSheet(
                f"color: {C_MUTED}; background: transparent; font-size: 10px;")
            italic_lbl.setCursor(Qt.PointingHandCursor)
            italic_lbl.mousePressEvent = lambda _e, b=italic_btn: b.toggle()
            ital_lay.addWidget(italic_lbl)
        ital_lay.addWidget(italic_btn)
        top.addWidget(ital_wrap)

        color_wrap = QWidget()
        color_wrap.setStyleSheet("background: transparent;")
        color_lay = QHBoxLayout(color_wrap)
        color_lay.setContentsMargins(0, 0, 0, 0); color_lay.setSpacing(6)
        if show_labels:
            color_lbl = QLabel(tr(T.color))
            color_lbl.setStyleSheet(
                f"color: {C_MUTED}; background: transparent; font-size: 10px;")
            color_lay.addWidget(color_lbl)
        color_btn = QPushButton()
        color_btn.setFixedSize(16, 16)
        color_btn.setCursor(Qt.PointingHandCursor)
        color_btn.setToolTip(tr(T.color_tooltip))
        color_btn.setProperty("color", "#000000")
        apply_color_button_style(color_btn, "#000000")
        color_btn.clicked.connect(lambda _=False, k=key: self._pick_color(k))
        color_lay.addWidget(color_btn)
        top.addWidget(color_wrap)

        bl.addLayout(top)

        inp = QLineEdit()
        inp.setFixedHeight(42)
        inp.setStyleSheet(f"""
            QLineEdit {{
                background-color: #FFFFFF;
                border: 1.5px solid {C_BORDER};
                border-radius: 10px;
                padding: 0 14px;
                color: {C_DARK_GREEN};
                font-size: 11px;
            }}
            QLineEdit:focus {{ border-color: {C_GREEN}; }}
        """)
        bl.addWidget(inp)

        self._fields[key] = {
            "input": inp,
            "size": size_combo,
            "bold": bold_btn,
            "italic": italic_btn,
            "color": color_btn,
        }
        return box

    def _on_confirm(self):
        if not HAS_DOCX:
            ThemedDialog.error(self, tr(T.missing_library), tr(T.no_docx))
            return
        if not self._template:
            return

        filename = self.filename_input.text().strip()
        if not filename:
            ThemedDialog.warning(self, tr(T.error), tr(T.err_empty_filename))
            return

        if self._fields:
            any_filled = any(
                w["input"].text().strip() for w in self._fields.values()
            )
            if not any_filled:
                ThemedDialog.warning(self, tr(T.error), tr(T.err_fill_one))
                return

        for ch in '<>:"/\\|?*':
            filename = filename.replace(ch, "_")
        filename = filename.strip()
        if not filename.lower().endswith(".docx"):
            filename += ".docx"

        output_dir = Path(SETTINGS.get("output_dir", str(Path.home() / "Desktop")))
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            ThemedDialog.error(self, tr(T.error),
                               tr(T.err_output_dir_create, err=e))
            return

        out_path = output_dir / filename
        src = Path(self._template.get("address", ""))
        if not src.exists():
            ThemedDialog.error(self, tr(T.error),
                               tr(T.err_template_not_found, path=src))
            return

        if out_path.exists():
            ok = ThemedDialog.confirm(
                self, tr(T.file_exists),
                tr(T.file_overwrite, f=filename))
            if not ok: return

        replacements = {}
        font_overrides = {}
        for key, widgets in self._fields.items():
            replacements[key] = widgets["input"].text()

            size_text = widgets["size"].currentText().strip().lower()
            if size_text in ("", "auto", tr(T.auto).strip().lower()):
                size_val = None
            else:
                try:
                    digit_map = str.maketrans(
                        "۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩",
                        "01234567890123456789",
                    )
                    normalized_size = (
                        size_text.translate(digit_map)
                        .replace(",", ".")
                        .replace("٫", ".")
                    )
                    size_val = float(normalized_size)
                except ValueError:
                    size_val = None
            bold_val = widgets["bold"].isChecked()
            italic_val = widgets["italic"].isChecked()
            color_val = widgets["color"].property("color") or "#000000"
            font_overrides[key] = (size_val, bold_val, italic_val, color_val)

        try:
            replace_placeholders_in_docx(
                str(src), str(out_path), replacements, font_overrides)
            ThemedDialog.info(
                self, tr(T.issued),
                tr(T.issued_ok, path=out_path))
            self.back_requested.emit()
        except Exception as ex:
            ThemedDialog.error(self, tr(T.error), tr(T.err_issue, err=ex))
            print("[error]", ex)


class BulkIssuePage(QWidget):
    back_requested = Signal()

    def __init__(self):
        super().__init__()
        self._template = None
        self.variable_names = []
        self._font_options = {}
        self.setStyleSheet("background: transparent;")

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        header = QWidget()
        header.setStyleSheet("background: transparent;")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(40, 18, 40, 12)
        header_layout.setSpacing(10)

        title = QLabel(tr(T.batch_issue_title))
        title_font = QFont(); title_font.setPointSize(15); title_font.setBold(True)
        title.setFont(title_font)
        title.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")

        back_btn = QPushButton(tr(T.back_spaced))
        back_btn.setIcon(icon_arrow_left(16))
        back_btn.setIconSize(QSize(16, 16))
        back_btn.setCursor(Qt.PointingHandCursor)
        back_btn.setFixedHeight(34)
        back_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent; border: 1px solid {C_BORDER};
                border-radius: 8px; color: {C_DARK_GREEN};
                padding: 0 14px; font-size: 11px;
            }}
            QPushButton:hover {{
                background-color: {C_LIGHT_GREEN}; border-color: {C_GREEN};
            }}
        """)
        back_btn.clicked.connect(self.back_requested.emit)
        header_layout.addWidget(title)
        header_layout.addStretch()
        header_layout.addWidget(back_btn)
        root.addWidget(header)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.scroll.setStyleSheet(f"""
            QScrollArea {{ background: transparent; border: none; }}
            QScrollBar:vertical {{
                background: transparent; width: 8px;
            }}
            QScrollBar::handle:vertical {{
                background: {C_BORDER}; border-radius: 4px; min-height: 30px;
            }}
            QScrollBar::handle:vertical:hover {{ background: {C_GREEN}; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0;
            }}
        """)

        body = QWidget()
        body.setStyleSheet("background: transparent;")
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(40, 0, 40, 18)
        body_layout.setSpacing(10)

        self.explanation_lbl = QLabel(tr(T.batch_explanation))
        self.explanation_lbl.setWordWrap(True)
        self.explanation_lbl.setStyleSheet(
            f"color: {C_DARK_GREEN}; background: {C_LIGHT_GREEN}; "
            f"border: 1px solid {C_BORDER}; border-radius: 10px; padding: 10px;"
        )
        body_layout.addWidget(self.explanation_lbl)

        self.mapping_title = QLabel(tr(T.batch_columns_title))
        self.mapping_title.setStyleSheet(
            f"color: {C_DARK_GREEN}; background: transparent; font-weight: bold;"
        )
        body_layout.addWidget(self.mapping_title)

        self.mapping_container = QWidget()
        self.mapping_container.setStyleSheet(
            f"background: #FFFFFF; border: 1px solid {C_BORDER}; border-radius: 8px;"
        )
        self.mapping_layout = QVBoxLayout(self.mapping_container)
        self.mapping_layout.setContentsMargins(10, 8, 10, 8)
        self.mapping_layout.setSpacing(4)
        body_layout.addWidget(self.mapping_container)

        self.header_row_check = QCheckBox(tr(T.batch_header_row))
        self.header_row_check.setCursor(Qt.PointingHandCursor)
        self.header_row_check.setStyleSheet(
            f"QCheckBox {{ color: {C_DARK_GREEN}; background: transparent; spacing: 8px; }}"
            f"QCheckBox::indicator {{ width: 16px; height: 16px; }}"
        )
        body_layout.addWidget(self.header_row_check)

        self.format_title = QLabel(tr(T.batch_format_title))
        self.format_title.setStyleSheet(
            f"color: {C_DARK_GREEN}; background: transparent; font-weight: bold;"
        )
        body_layout.addWidget(self.format_title)

        self.options_container = QWidget()
        self.options_container.setStyleSheet("background: transparent;")
        self.options_layout = QVBoxLayout(self.options_container)
        self.options_layout.setContentsMargins(0, 0, 0, 0)
        self.options_layout.setSpacing(6)
        body_layout.addWidget(self.options_container)

        self.excel_zone = ExcelDropZone()
        body_layout.addWidget(self.excel_zone)
        self.scroll.setWidget(body)
        root.addWidget(self.scroll, 1)

        footer = QWidget()
        footer.setStyleSheet("background: transparent;")
        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(40, 8, 40, 18)
        self.confirm_btn = QPushButton(tr(T.batch_confirm))
        self.confirm_btn.setCursor(Qt.PointingHandCursor)
        self.confirm_btn.setFixedHeight(44)
        self.confirm_btn.setMinimumWidth(200)
        self.confirm_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {C_GREEN}; color: white;
                border: none; border-radius: 10px;
                font-size: 12px; font-weight: bold; padding: 0 24px;
            }}
            QPushButton:hover {{ background-color: {C_DARK_GREEN}; }}
            QPushButton:disabled {{ background-color: {C_MUTED}; }}
        """)
        self.confirm_btn.clicked.connect(self._on_confirm)
        footer_layout.addWidget(self.confirm_btn)
        footer_layout.addStretch()
        root.addWidget(footer)
        self.confirm_btn.setEnabled(False)

    def _clear_layout(self, layout):
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

    def load_template(self, template):
        self._template = template or {}
        self.variable_names = []
        self._font_options = {}
        self.header_row_check.setChecked(False)
        self.excel_zone.reset()
        self._clear_layout(self.mapping_layout)
        self._clear_layout(self.options_layout)

        source = Path(self._template.get("address", ""))
        if not source.is_file():
            ThemedDialog.warning(
                self, tr(T.error), tr(T.err_template_not_found, path=source)
            )
            self.confirm_btn.setEnabled(False)
            return

        try:
            properties = extract_placeholders(str(source))
        except WrapperCountError as exc:
            ThemedDialog.error(self, tr(T.err_wrapper_count), str(exc))
            self.confirm_btn.setEnabled(False)
            return
        except Exception as exc:
            ThemedDialog.error(self, tr(T.error), tr(T.cli_read_error, err=exc))
            self.confirm_btn.setEnabled(False)
            return

        self.variable_names = list(properties.keys())
        if not self.variable_names:
            ThemedDialog.warning(
                self, tr(T.error), tr(T.cli_no_variables, w=_get_wrapper())
            )
            self.confirm_btn.setEnabled(False)
            return

        self.mapping_layout.addWidget(QLabel(
            tr(T.batch_mapping_filename, column=1)
        ))
        for index, variable in enumerate(self.variable_names, start=2):
            label = QLabel(tr(
                T.batch_mapping_variable,
                column=index,
                name=variable,
            ))
            label.setWordWrap(True)
            self.mapping_layout.addWidget(label)

        for index, variable in enumerate(self.variable_names, start=2):
            card = QFrame()
            card.setStyleSheet(f"""
                QFrame {{
                    background: #FFFFFF; border: 1px solid {C_BORDER};
                    border-radius: 9px;
                }}
            """)
            row = QHBoxLayout(card)
            row.setContentsMargins(10, 6, 10, 6)
            row.setSpacing(8)

            variable_label = QLabel(tr(
                T.batch_mapping_variable,
                column=index,
                name=variable,
            ))
            variable_label.setWordWrap(True)
            variable_label.setStyleSheet(
                f"color: {C_DARK_GREEN}; background: transparent; border: none;"
            )
            row.addWidget(variable_label, 1)

            size_combo = FontSizeCombo()
            size_combo.setFixedWidth(94)
            size_combo.setToolTip(tr(T.font_size))
            row.addWidget(size_combo)

            bold_toggle = RoundToggle("B", tr(T.bold), bold=True)
            row.addWidget(bold_toggle)
            italic_toggle = RoundToggle("I", tr(T.italic), italic=True)
            row.addWidget(italic_toggle)

            color_btn = QPushButton()
            color_btn.setFixedSize(20, 20)
            color_btn.setCursor(Qt.PointingHandCursor)
            color_btn.setToolTip(tr(T.color_tooltip))
            color_btn.setProperty("color", "#000000")
            apply_color_button_style(color_btn, "#000000")
            color_btn.clicked.connect(
                lambda _=False, key=variable: self._pick_color(key)
            )
            row.addWidget(color_btn)

            self.options_layout.addWidget(card)
            self._font_options[variable] = {
                "size": size_combo,
                "bold": bold_toggle,
                "italic": italic_toggle,
                "color": color_btn,
            }

        self.confirm_btn.setEnabled(True)

    def _pick_color(self, key):
        controls = self._font_options.get(key)
        if not controls:
            return
        color_btn = controls["color"]
        current = color_btn.property("color") or "#000000"
        dialog = QColorDialog(QColor(current), self)
        dialog.setOption(QColorDialog.ColorDialogOption.DontUseNativeDialog, True)
        dialog.setOption(QColorDialog.ColorDialogOption.ShowAlphaChannel, False)
        dialog.setLayoutDirection(Qt.LeftToRight)
        dialog.setWindowTitle(tr(T.pick_color))

        palette = dialog.palette()
        palette.setColor(QPalette.ColorRole.Window, QColor("#FFFFFF"))
        palette.setColor(QPalette.ColorRole.WindowText, QColor(C_DARK_GREEN))
        palette.setColor(QPalette.ColorRole.Base, QColor("#FFFFFF"))
        palette.setColor(QPalette.ColorRole.Text, QColor(C_DARK_GREEN))
        palette.setColor(QPalette.ColorRole.Button, QColor(C_LIGHT_GREEN))
        palette.setColor(QPalette.ColorRole.ButtonText, QColor(C_DARK_GREEN))
        palette.setColor(QPalette.ColorRole.Highlight, QColor(C_GREEN))
        palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#FFFFFF"))
        dialog.setPalette(palette)
        dialog.setStyleSheet(f"""
            QColorDialog {{ background: #FFFFFF; color: {C_DARK_GREEN}; }}
            QColorDialog QLabel {{ color: {C_DARK_GREEN}; background: transparent; }}
            QColorDialog QPushButton {{
                color: {C_DARK_GREEN}; background: {C_LIGHT_GREEN};
                border: 1px solid {C_BORDER}; border-radius: 6px; padding: 5px 12px;
            }}
        """)

        def localize_hex_label():
            dialog.setLayoutDirection(Qt.LeftToRight)
            for child in dialog.findChildren(QWidget):
                child.setLayoutDirection(Qt.LeftToRight)
            for label in dialog.findChildren(QLabel):
                text = label.text().replace("&", "").strip().rstrip(":").strip().casefold()
                if text == "html":
                    label.setText("HEX:")
                    break

        QTimer.singleShot(0, localize_hex_label)
        if dialog.exec():
            selected = dialog.currentColor()
            if selected.isValid():
                color_hex = selected.name()
                color_btn.setProperty("color", color_hex)
                apply_color_button_style(color_btn, color_hex)

    def _collect_font_overrides(self):
        result = {}
        for variable, controls in self._font_options.items():
            result[variable] = (
                _font_size_value(controls["size"]),
                controls["bold"].isChecked(),
                controls["italic"].isChecked(),
                controls["color"].property("color") or "#000000",
            )
        return result

    def _on_confirm(self):
        if not HAS_DOCX:
            ThemedDialog.error(self, tr(T.missing_library), tr(T.no_docx))
            return
        if not HAS_EXCEL:
            ThemedDialog.error(
                self, tr(T.missing_library), tr(T.batch_excel_library_missing)
            )
            return
        if not self._template or not self.variable_names:
            return

        excel_path = self.excel_zone.file_path()
        if not excel_path:
            ThemedDialog.warning(self, tr(T.error), tr(T.batch_excel_missing))
            return

        source = Path(self._template.get("address", ""))
        if not source.is_file():
            ThemedDialog.warning(
                self, tr(T.error), tr(T.err_template_not_found, path=source)
            )
            return

        try:
            records = read_excel_records(
                excel_path,
                self.variable_names,
                header_row=self.header_row_check.isChecked(),
            )
        except ExcelImportError as exc:
            ThemedDialog.error(self, tr(T.error), str(exc))
            return

        configured_output = str(SETTINGS.get("output_dir") or "").strip()
        base_output = Path(
            configured_output or (Path.home() / "Desktop")
        ).expanduser()
        batch_directory = base_output / tr(T.batch_output_folder)
        try:
            batch_directory.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            ThemedDialog.error(
                self, tr(T.error), tr(T.batch_output_folder_error, err=exc)
            )
            return

        reserved_names = set()
        font_overrides = self._collect_font_overrides()
        completed = 0
        for row_number, output_name, replacements in records:
            output_file = _safe_batch_filename(
                output_name, batch_directory, reserved_names
            )
            try:
                _save_docx_atomically(
                    source,
                    batch_directory / output_file,
                    replacements,
                    font_overrides,
                    overwrite=False,
                )
                completed += 1
            except Exception as exc:
                ThemedDialog.error(
                    self,
                    tr(T.error),
                    tr(T.batch_row_write_error, row=row_number, err=exc),
                )
                return

        ThemedDialog.info(
            self,
            tr(T.issued),
            tr(T.batch_success, count=completed, path=batch_directory),
        )


# ═════════════════════════════════════════════════════════════════
#  SettingsPage — no theme option
# ═════════════════════════════════════════════════════════════════

class SettingsPage(QWidget):
    back_requested = Signal()
    saved = Signal()

    IN_H      = 36
    FIELD_W   = 380
    BROWSE_W  = 80

    LANGUAGES = [
        ("fa", "فارسی"),
        ("en", "English"),
    ]

    def __init__(self):
        super().__init__()
        self.setStyleSheet("background: transparent;")

        lay = QVBoxLayout(self)
        lay.setContentsMargins(40, 22, 40, 30); lay.setSpacing(10)

        header = QHBoxLayout()
        title = QLabel(tr(T.settings))
        tf = QFont(); tf.setPointSize(15); tf.setBold(True)
        title.setFont(tf)
        title.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")

        back_btn = QPushButton(tr(T.back_spaced))
        back_btn.setIcon(icon_arrow_left(16)); back_btn.setIconSize(QSize(16, 16))
        back_btn.setCursor(Qt.PointingHandCursor); back_btn.setFixedHeight(34)
        back_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent; border: 1px solid {C_BORDER};
                border-radius: 8px; color: {C_DARK_GREEN};
                padding: 0 14px; font-size: 11px;
            }}
            QPushButton:hover {{
                background-color: {C_LIGHT_GREEN};
                border-color: {C_GREEN};
            }}
        """)
        back_btn.clicked.connect(self.back_requested.emit)

        header.addWidget(title); header.addStretch(); header.addWidget(back_btn)
        lay.addLayout(header)

        lf = QFont(); lf.setPointSize(10); lf.setBold(True)

        lang_lbl = QLabel(tr(T.app_language))
        lang_lbl.setFont(lf)
        lang_lbl.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")
        lay.addWidget(lang_lbl)

        self.lang_combo = ComboWithArrow(editable=False)
        self.lang_combo.setFixedHeight(self.IN_H)
        self.lang_combo.setFixedWidth(self.FIELD_W)
        for code, label in self.LANGUAGES:
            self.lang_combo.addItem(label, code)
        cur_lang = SETTINGS.get("language", "fa")
        idx_lang = self.lang_combo.findData(cur_lang)
        if idx_lang >= 0:
            self.lang_combo.setCurrentIndex(idx_lang)
        self._style_combo(self.lang_combo)

        lang_row = QHBoxLayout()
        lang_row.setContentsMargins(0, 0, 0, 0)
        lang_row.addWidget(self.lang_combo)
        lang_row.addStretch()
        lay.addLayout(lang_row); lay.addSpacing(2)

        out_lbl = QLabel(tr(T.output_path))
        out_lbl.setFont(lf)
        out_lbl.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")
        lay.addWidget(out_lbl)

        out_container = QWidget()
        out_container.setStyleSheet("background: transparent;")
        out_container.setFixedWidth(self.FIELD_W)
        out_row = QHBoxLayout(out_container)
        out_row.setContentsMargins(0, 0, 0, 0); out_row.setSpacing(8)

        self.output_input = QLineEdit()
        self.output_input.setText(SETTINGS.get("output_dir", str(Path.home() / "Desktop")))
        self.output_input.setFixedHeight(self.IN_H)
        self.output_input.setStyleSheet(f"""
            QLineEdit {{
                background-color: #FFFFFF;
                border: 1.5px solid {C_BORDER};
                border-radius: 8px;
                padding: 0 12px;
                color: {C_DARK_GREEN};
                font-size: 11px;
            }}
            QLineEdit:focus {{ border-color: {C_GREEN}; }}
        """)
        out_row.addWidget(self.output_input, 1)

        browse_btn = QPushButton(tr(T.browse))
        browse_btn.setFixedHeight(self.IN_H)
        browse_btn.setFixedWidth(self.BROWSE_W)
        browse_btn.setCursor(Qt.PointingHandCursor)
        browse_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {C_LIGHT_GREEN};
                border: 1px solid {C_BORDER};
                border-radius: 8px;
                color: {C_DARK_GREEN};
                font-size: 11px;
            }}
            QPushButton:hover {{
                background-color: {C_HOVER};
                border-color: {C_GREEN};
            }}
        """)
        browse_btn.clicked.connect(self._pick_output_dir)
        out_row.addWidget(browse_btn)

        out_wrap = QHBoxLayout()
        out_wrap.setContentsMargins(0, 0, 0, 0)
        out_wrap.addWidget(out_container)
        out_wrap.addStretch()
        lay.addLayout(out_wrap); lay.addSpacing(2)

        wrap_lbl = QLabel(tr(T.placeholder_wrapper))
        wrap_lbl.setFont(lf)
        wrap_lbl.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")
        lay.addWidget(wrap_lbl)

        self.wrapper_input = QLineEdit()
        self.wrapper_input.setText(SETTINGS.get("placeholder_wrapper", '##'))
        self.wrapper_input.setFixedHeight(self.IN_H)
        self.wrapper_input.setFixedWidth(self.FIELD_W)
        self.wrapper_input.setPlaceholderText(tr(T.ph_wrapper))
        self.wrapper_input.setStyleSheet(f"""
            QLineEdit {{
                background-color: #FFFFFF;
                border: 1.5px solid {C_BORDER};
                border-radius: 8px;
                padding: 0 12px;
                color: {C_DARK_GREEN};
                font-size: 12px;
                font-family: Consolas, monospace;
            }}
            QLineEdit:focus {{ border-color: {C_GREEN}; }}
        """)

        wrap_row = QHBoxLayout()
        wrap_row.setContentsMargins(0, 0, 0, 0)
        wrap_row.addWidget(self.wrapper_input)
        wrap_row.addStretch()
        lay.addLayout(wrap_row)

        example = self.wrapper_input.text() + tr(T.user_name) + self.wrapper_input.text()
        hint = QLabel(tr(T.wrapper_hint, example=example))
        hint.setStyleSheet(f"color: {C_MUTED}; background: transparent; font-size: 10px;")
        hint.setWordWrap(True)
        lay.addWidget(hint)

        lay.addStretch()

        btn_row = QHBoxLayout()
        self.save_btn = QPushButton(tr(T.save_settings))
        self.save_btn.setCursor(Qt.PointingHandCursor)
        self.save_btn.setFixedHeight(42); self.save_btn.setMinimumWidth(160)
        self.save_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {C_GREEN}; color: white;
                border: none; border-radius: 10px;
                font-size: 12px; font-weight: bold;
                padding: 0 24px;
            }}
            QPushButton:hover {{ background-color: {C_DARK_GREEN}; }}
        """)
        self.save_btn.clicked.connect(self._on_save)
        btn_row.addWidget(self.save_btn)
        btn_row.addStretch()
        lay.addLayout(btn_row)

    def _style_combo(self, combo):
        combo.setStyleSheet(f"""
            QComboBox {{
                background-color: #FFFFFF;
                border: 1.5px solid {C_BORDER};
                border-radius: 8px;
                padding-left: 26px;
                padding-right: 10px;
                color: {C_DARK_GREEN};
                font-size: 11px;
            }}
            QComboBox:focus, QComboBox:on {{ border-color: {C_GREEN}; }}
            QComboBox QAbstractItemView {{
                background-color: #FFFFFF;
                border: 1px solid {C_BORDER};
                border-radius: 8px;
                padding: 4px;
                outline: none;
                color: {C_DARK_GREEN};
                selection-background-color: {C_LIGHT_GREEN};
                selection-color: {C_DARK_GREEN};
            }}
            QComboBox QAbstractItemView::item {{
                padding: 6px; border-radius: 6px; min-height: 20px;
            }}
            QComboBox QAbstractItemView::item:hover {{
                background-color: {C_HOVER};
            }}
        """)

    def _pick_output_dir(self):
        start = self.output_input.text() or str(Path.home())
        d = QFileDialog.getExistingDirectory(self, tr(T.select_output_dir), start)
        if d:
            self.output_input.setText(d)

    def _on_save(self):
        new_lang = self.lang_combo.currentData()
        new_output = self.output_input.text().strip()
        new_wrapper = self.wrapper_input.text()

        if not new_output:
            ThemedDialog.warning(self, tr(T.error), tr(T.err_output_empty))
            return
        try:
            Path(new_output).mkdir(parents=True, exist_ok=True)
        except Exception as e:
            ThemedDialog.error(self, tr(T.error),
                               tr(T.err_output_invalid, err=e))
            return

        if len(new_wrapper) == 0:
            ThemedDialog.warning(self, tr(T.error), tr(T.err_wrapper_empty))
            return

        SETTINGS["language"] = new_lang
        SETTINGS["output_dir"] = new_output
        SETTINGS["placeholder_wrapper"] = new_wrapper
        save_settings()

        self.saved.emit()


# ═════════════════════════════════════════════════════════════════
#  Container — plain background, no decorations
# ═════════════════════════════════════════════════════════════════

class Container(QWidget):
    def __init__(self):
        super().__init__()
        self.setStyleSheet("background: transparent;")

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        w, h = self.width(), self.height()

        try:
            is_max = self.window().isMaximized()
        except Exception:
            is_max = False
        radius = 0 if is_max else 14

        path = QPainterPath()
        path.addRoundedRect(QRectF(0.5, 0.5, w - 1, h - 1), radius, radius)
        p.fillPath(path, QColor(C_BG))

        if radius > 0:
            p.setPen(QPen(QColor(C_BORDER), 1))
            p.setBrush(Qt.NoBrush)
            p.drawPath(path)


# ═════════════════════════════════════════════════════════════════
#  MainWindow
# ═════════════════════════════════════════════════════════════════

class MainWindow(QMainWindow):
    theme_changed = Signal()
    RESIZE_MARGIN = 6

    def __init__(self):
        super().__init__()
        self._built_language = SETTINGS.get("language", "fa")
        self._really_quit = False
        self.setWindowTitle(tr(T.app_title))
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Window)
        self.setAttribute(Qt.WA_TranslucentBackground)
        saved_size = SETTINGS.get("window_size", [720, 600])
        try:
            if not isinstance(saved_size, (list, tuple)) or len(saved_size) < 2:
                raise ValueError("invalid saved window size")
            start_w = max(300, int(saved_size[0]))
            start_h = max(300, int(saved_size[1]))
        except (TypeError, ValueError, IndexError):
            start_w, start_h = 720, 600
        screen = QGuiApplication.primaryScreen()
        if screen is not None:
            available = screen.availableGeometry()
            start_w = min(start_w, max(300, available.width()))
            start_h = min(start_h, max(300, available.height()))
        self.resize(start_w, start_h)
        self._window_state_timer = QTimer(self)
        self._window_state_timer.setSingleShot(True)
        self._window_state_timer.timeout.connect(self._save_window_state)
        self.setMinimumSize(300, 300)

        self.setWindowFlags(Qt.Window | Qt.FramelessWindowHint)

        root = QWidget()
        root.setAttribute(Qt.WA_TranslucentBackground)
        self._root_layout = QVBoxLayout(root)
        self._root_layout.setContentsMargins(0, 0, 0, 0)

        self.container = Container()
        self._root_layout.addWidget(self.container)
        self.setCentralWidget(root)

        v = QVBoxLayout(self.container)
        v.setContentsMargins(0, 0, 0, 0); v.setSpacing(0)
        self.title_bar = TitleBar(self)
        v.addWidget(self.title_bar)

        self.stack = QStackedWidget()
        self.stack.setStyleSheet("background: transparent;")
        v.addWidget(self.stack, 1)

        self.home_page = HomePage()
        self.add_page = AddLetterPage()
        self.edit_page = EditLetterPage()
        self.bulk_page = BulkIssuePage()
        self.settings_page = SettingsPage()

        for w in (
            self.home_page, self.add_page, self.edit_page,
            self.bulk_page, self.settings_page,
        ):
            self.stack.addWidget(w)

        self.home_page.dropdown.add_requested.connect(self._go_to_add)
        self.home_page.dropdown.template_selected.connect(self._on_template_selected)
        self.home_page.dropdown.template_deleted.connect(self._on_template_deleted)
        self.home_page.settings_requested.connect(self._go_to_settings)
        self.home_page.quick_template_selected.connect(self._on_template_selected)
        self.home_page.quick_access_changed.connect(self._refresh_tray_menu)
        
        self.add_page.back_requested.connect(self._go_to_home)
        self.add_page.saved.connect(self._on_letter_saved)

        self.edit_page.back_requested.connect(self._go_to_home)
        self.edit_page.batch_requested.connect(self._open_bulk_issue)
        self.bulk_page.back_requested.connect(self._return_to_issue_page)

        self.settings_page.back_requested.connect(self._go_to_home)
        self.settings_page.saved.connect(self._on_settings_saved)

        self.reload_templates()
        self.reload_templates()
        self._setup_tray()

    # ═════════════════════════════════════════════════════════════
    #  System Tray
    # ═════════════════════════════════════════════════════════════

    def _setup_tray(self):
        self.tray_icon = QSystemTrayIcon(self)
        icon_path = APP_ROOT / "app.ico"
        if icon_path.exists():
            self.tray_icon.setIcon(QIcon(str(icon_path)))
        else:
            self.tray_icon.setIcon(icon_document(32, C_DARK_GREEN))
        self.tray_icon.setToolTip(tr(T.app_title))

        self.tray_menu = QMenu()
        self._build_tray_menu()

        self.tray_icon.setContextMenu(self.tray_menu)
        self.tray_icon.activated.connect(self._on_tray_activated)
        self.tray_icon.show()

    def _build_tray_menu(self):
        self.tray_menu.clear()

        open_action = QAction(
            tr({"persian": "باز کردن برنامه", "english": "Open App"}), self)
        open_action.triggered.connect(self._show_from_tray)
        self.tray_menu.addAction(open_action)

        qa_names = [n for n in SETTINGS.get("quick_access", []) if n]
        if qa_names:
            templates = load_templates()
            tpl_map = {t.get("name", ""): t for t in templates}
            valid_qa = [n for n in qa_names if n in tpl_map]

            if valid_qa:
                self.tray_menu.addSeparator()

                header = QAction(tr(T.quick_access), self)
                header.setEnabled(False)
                self.tray_menu.addAction(header)

                for name in valid_qa:
                    act = QAction(name, self)
                    act.triggered.connect(
                        lambda _=None, n=name:
                            self._open_template_from_tray(n))
                    self.tray_menu.addAction(act)

        self.tray_menu.addSeparator()
        exit_action = QAction(
            tr({"persian": "خروج", "english": "Exit"}), self)
        exit_action.triggered.connect(self._quit_app)
        self.tray_menu.addAction(exit_action)

    def _refresh_tray_menu(self):
        if hasattr(self, "tray_menu"):
            self._build_tray_menu()

    def _show_from_tray(self):
        self.show()
        if self.isMinimized():
            self.showNormal()
        self.raise_()
        self.activateWindow()

    def _open_template_from_tray(self, name):
        templates = load_templates()
        tpl_map = {t.get("name", ""): t for t in templates}
        tpl = tpl_map.get(name)
        if tpl is None:
            self._refresh_tray_menu()
            return
        self._show_from_tray()
        self._on_template_selected(tpl)

    def _on_tray_activated(self, reason):
        if reason in (QSystemTrayIcon.Trigger, QSystemTrayIcon.DoubleClick):
            if self.isVisible() and not self.isMinimized():
                self.hide()
            else:
                self._show_from_tray()

    def _quit_app(self):
        self._really_quit = True
        QApplication.instance().quit()

    def _apply_win32_styles(self):
        """استایل WS_THICKFRAME و WS_CAPTION را در سطح Win32 فعال می‌کند
        تا ویندوز پنجره را به عنوان یک پنجره قابل Snap و Resize بشناسد."""
        if sys.platform != "win32":
            return
        try:
            from win32gui import GetWindowLong, SetWindowLong
            from win32con import WS_THICKFRAME, WS_CAPTION, WS_SYSMENU, WS_MINIMIZEBOX, WS_MAXIMIZEBOX, GWL_STYLE

            hwnd = int(self.winId())

            style = GetWindowLong(hwnd, GWL_STYLE)
            style |= WS_THICKFRAME      
            style |= WS_CAPTION         
            style |= WS_SYSMENU         
            style |= WS_MINIMIZEBOX     
            style |= WS_MAXIMIZEBOX     
            SetWindowLong(hwnd, GWL_STYLE, style)
        except Exception as e:
            print(f"[Aero Snap] init styles failed: {e}")

    def nativeEvent(self, eventType, message):
        if sys.platform == "win32":
            try:
                import ctypes
                from ctypes import wintypes
                try:
                    addr = int(message)
                except Exception:
                    addr = ctypes.cast(message, ctypes.c_void_p).value
                msg = wintypes.MSG.from_address(addr)

                WM_NCCALCSIZE = 0x0083
                WM_NCHITTEST = 0x0084

                if msg.message == WM_NCCALCSIZE:
                    return True, 0

                if msg.message == WM_NCHITTEST and not self.isMaximized():
                    x = ctypes.c_short(msg.lParam & 0xFFFF).value
                    y = ctypes.c_short((msg.lParam >> 16) & 0xFFFF).value

                    dpr = 1.0
                    try:
                        wh = self.windowHandle()
                        if wh is not None:
                            dpr = wh.devicePixelRatio() or 1.0
                    except Exception:
                        pass

                    local = self.mapFromGlobal(QPoint(int(x / dpr), int(y / dpr)))
                    w, h = self.width(), self.height()
                    m = self.RESIZE_MARGIN

                    on_l = 0 <= local.x() < m
                    on_r = w - m < local.x() <= w
                    on_t = 0 <= local.y() < m
                    on_b = h - m < local.y() <= h

                    if on_t and on_l: return True, 13  # HTTOPLEFT
                    if on_t and on_r: return True, 14  # HTTOPRIGHT
                    if on_b and on_l: return True, 16  # HTBOTTOMLEFT
                    if on_b and on_r: return True, 17  # HTBOTTOMRIGHT
                    if on_l:          return True, 10  # HTLEFT
                    if on_r:          return True, 11  # HTRIGHT
                    if on_t:          return True, 12  # HTTOP
                    if on_b:          return True, 15  # HTBOTTOM

            except Exception:
                pass

        return super().nativeEvent(eventType, message)

    def showEvent(self, e):
        super().showEvent(e)
        self._apply_win32_styles()
        
    def reload_templates(self):
        self.home_page.set_templates(load_templates())
        if hasattr(self, "tray_menu"):
            self._refresh_tray_menu()

    def _go_to_home(self):
        self.stack.setCurrentWidget(self.home_page)

    def _open_bulk_issue(self):
        template = self.edit_page._template
        if not template:
            return
        self.bulk_page.load_template(template)
        self.stack.setCurrentWidget(self.bulk_page)

    def _return_to_issue_page(self):
        self.stack.setCurrentWidget(self.edit_page)

    def _go_to_add(self):
        self.add_page.reset()
        self.stack.setCurrentWidget(self.add_page)

    def _go_to_settings(self):
        self.stack.removeWidget(self.settings_page)
        self.settings_page.deleteLater()
        self.settings_page = SettingsPage()
        self.settings_page.back_requested.connect(self._go_to_home)
        self.settings_page.saved.connect(self._on_settings_saved)
        self.stack.addWidget(self.settings_page)
        self.stack.setCurrentWidget(self.settings_page)

    def _on_settings_saved(self):
        self._refresh_tray_menu()
        lang_changed = SETTINGS.get("language", "fa") != self._built_language
        if lang_changed:
            self.theme_changed.emit()
        else:
            self._go_to_home()

    def _on_letter_saved(self):
        self.reload_templates()
        self._go_to_home()

    def _on_template_selected(self, tpl):
        if not HAS_DOCX:
            ThemedDialog.error(self, tr(T.missing_library), tr(T.no_docx))
            return
        self.edit_page.load_template(tpl)
        self.stack.setCurrentWidget(self.edit_page)

    def _on_template_deleted(self, tpl):
        name = tpl.get("name", "")
        address = tpl.get("address", "")

        ok = ThemedDialog.confirm_delete(
            self, tr(T.delete_template),
            tr(T.confirm_delete_tpl, name=name)
        )
        if not ok: return

        data = load_templates_raw()
        others = [t for t in data
                  if t.get("name") != name and t.get("address") == address]

        data = [t for t in data if t.get("name") != name]
        save_templates(data)

        if not others and address:
            p = Path(address)
            if p.exists():
                try:
                    p.unlink()
                except Exception as e:
                    print(f"[error] file removal: {e}")

        qa = [n for n in SETTINGS.get("quick_access", []) if n != name]
        if qa != SETTINGS.get("quick_access", []):
            SETTINGS["quick_access"] = qa
            save_settings()

        self.reload_templates()

        ThemedDialog.info(self, tr(T.deleted),
                          tr(T.deleted_ok, name=name))

    def _schedule_window_state_save(self):
        timer = getattr(self, "_window_state_timer", None)
        if timer is not None:
            timer.start(350)

    def _save_window_state(self):
        try:
            if self.isMinimized():
                save_settings()
                return
            if self.isMaximized():
                SETTINGS["window_maximized"] = True
                normal = self.normalGeometry()
                if normal.isValid():
                    SETTINGS["window_size"] = [
                        max(300, int(normal.width())),
                        max(300, int(normal.height())),
                    ]
            else:
                SETTINGS["window_maximized"] = False
                SETTINGS["window_size"] = [
                    max(300, int(self.width())),
                    max(300, int(self.height())),
                ]
            save_settings()
        except Exception as exc:
            print(f"[window] Could not save window state: {exc}")

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._schedule_window_state_save()

    def closeEvent(self, event):
        timer = getattr(self, "_window_state_timer", None)
        if timer is not None:
            timer.stop()
        self._save_window_state()

        if not getattr(self, "_really_quit", False):
            event.ignore()
            self.hide()
            return

        if hasattr(self, "tray_icon"):
            self.tray_icon.hide()
        super().closeEvent(event)

    def changeEvent(self, e):
        if e.type() == QEvent.WindowStateChange:
            try:
                self.title_bar.refresh_radius()
            except Exception:
                pass
            try:
                self.container.update()
            except Exception:
                pass
            if self.isMaximized():
                SETTINGS["window_maximized"] = True
            elif not self.isMinimized():
                SETTINGS["window_maximized"] = False
            self._schedule_window_state_save()
        super().changeEvent(e)


# ═════════════════════════════════════════════════════════════════
#  First-run Language Selection
# ═════════════════════════════════════════════════════════════════

class LangButton(QFrame):
    clicked = Signal()

    def __init__(self, primary, secondary, parent=None):
        super().__init__(parent)
        self.setObjectName("langbtn")
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumHeight(160)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.setSpacing(4)
        lay.addStretch()

        self._primary = QLabel(primary)
        pf = QFont(); pf.setPointSize(26); pf.setBold(True)
        self._primary.setFont(pf)
        self._primary.setAlignment(Qt.AlignCenter)
        self._primary.setAttribute(Qt.WA_TransparentForMouseEvents)

        self._secondary = QLabel(secondary)
        sf = QFont(); sf.setPointSize(11)
        self._secondary.setFont(sf)
        self._secondary.setAlignment(Qt.AlignCenter)
        self._secondary.setAttribute(Qt.WA_TransparentForMouseEvents)

        lay.addWidget(self._primary)
        lay.addWidget(self._secondary)
        lay.addStretch()

        self._apply_style(hover=False)

    def _apply_style(self, hover):
        if hover:
            self.setStyleSheet(f"""
                QFrame#langbtn {{
                    background-color: {C_GREEN};
                    border: 2px solid {C_DARK_GREEN};
                    border-radius: 16px;
                }}
            """)
            self._primary.setStyleSheet("color: white; background: transparent;")
            self._secondary.setStyleSheet(
                "color: #E6EEE8; background: transparent;")
        else:
            self.setStyleSheet(f"""
                QFrame#langbtn {{
                    background-color: {C_LIGHT_GREEN};
                    border: 2px solid {C_BORDER};
                    border-radius: 16px;
                }}
            """)
            self._primary.setStyleSheet(
                f"color: {C_DARK_GREEN}; background: transparent;")
            self._secondary.setStyleSheet(
                f"color: {C_MUTED}; background: transparent;")

    def enterEvent(self, e):
        self._apply_style(True)
        super().enterEvent(e)

    def leaveEvent(self, e):
        self._apply_style(False)
        super().leaveEvent(e)

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self.clicked.emit()


class LanguageSelectDialog(QDialog):

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(
            Qt.Dialog | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setModal(True)
        self.setLayoutDirection(Qt.LeftToRight)
        self.setFixedSize(460, 580)

        self.selected_language = None

        outer = QVBoxLayout(self)
        outer.setContentsMargins(18, 18, 18, 18)

        frame = QFrame()
        frame.setObjectName("langFrame")
        frame.setStyleSheet(f"""
            QFrame#langFrame {{
                background-color: #FFFFFF;
                border-radius: 20px;
                border: 1px solid {C_BORDER};
            }}
        """)
        sh = QGraphicsDropShadowEffect()
        sh.setBlurRadius(44)
        sh.setOffset(0, 12)
        sh.setColor(QColor(0, 0, 0, 95))
        frame.setGraphicsEffect(sh)
        outer.addWidget(frame)

        v = QVBoxLayout(frame)
        v.setContentsMargins(28, 30, 28, 28)
        v.setSpacing(14)

        title = QLabel("Choose your language  •  زبان خود را انتخاب کنید")
        tf = QFont(); tf.setPointSize(12); tf.setBold(True)
        title.setFont(tf)
        title.setAlignment(Qt.AlignCenter)
        title.setWordWrap(True)
        title.setStyleSheet(
            f"color: {C_DARK_GREEN}; background: transparent;")
        v.addWidget(title)

        v.addSpacing(6)

        persian_btn = LangButton("فارسی", "Persian")
        english_btn = LangButton("English", "انگلیسی")

        persian_btn.clicked.connect(lambda: self._choose("fa"))
        english_btn.clicked.connect(lambda: self._choose("en"))

        v.addWidget(persian_btn, 1)
        v.addWidget(english_btn, 1)

    def _choose(self, lang):
        self.selected_language = lang
        self.accept()

    def showEvent(self, e):
        super().showEvent(e)
        scr = QGuiApplication.primaryScreen()
        if scr is not None:
            geo = scr.availableGeometry()
            self.move(geo.center().x() - self.width() // 2,
                      geo.center().y() - self.height() // 2)

# ═════════════════════════════════════════════════════════════════
#  Command-line document issuance
# ═════════════════════════════════════════════════════════════════

def _build_cli_parser():
    """Create the parser used by both python and the frozen executable."""
    parser = argparse.ArgumentParser(
        prog=Path(sys.argv[0]).name or "WordCloner",
        description="Issue a WordCloner document from a template or Excel workbook.",
        epilog=(
            "Single document:\n"
            '  WordCloner.exe -i "letter_template.docx" -o "output.docx" '
            '-v "A. User" " " "Operations"\n\n'
            "Count template variables:\n"
            '  WordCloner.exe -i "letter_template.docx" --count\n\n'
            "Bulk from Excel:\n"
            '  WordCloner.exe -i "letter_template.docx" --excel "data.xlsx" '
            "--header-row"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "-i", "--input", dest="input_path", required=True, metavar="DOCX",
        help="Path to the .docx template.",
    )
    parser.add_argument(
        "-o", "--output", dest="output_path", default=None, metavar="PATH",
        help=(
            "Output .docx file for single issuance; output folder for --excel "
            "(default: settings output folder / Bulk Issuance)."
        ),
    )
    parser.add_argument(
        "-v", "--values", dest="values", nargs="*", default=None,
        metavar="VALUE",
        help=(
            'Values in first-appearance order. Pass a quoted single space (" ") '
            "for an intentionally blank variable."
        ),
    )
    parser.add_argument(
        "-w", "--wrapper", dest="wrapper", default=None, metavar="TEXT",
        help="Use TEXT as the placeholder wrapper for this run (default: settings.json).",
    )
    parser.add_argument(
        "-x", "--excel", dest="excel_path", default=None, metavar="XLSX",
        help="Read one output filename and the variable values from each Excel row.",
    )
    parser.add_argument(
        "--header-row", action="store_true",
        help="Skip row 1 of the active worksheet because it contains column headings.",
    )
    parser.add_argument(
        "-c", "--count", action="store_true",
        help="Print the number of unique variables in the template and exit.",
    )
    parser.add_argument(
        "--force", action="store_true",
        help="Replace an existing single-document output file.",
    )
    return parser


def _cli_output_path(raw_path, parser):
    path = Path(raw_path).expanduser()
    if path.exists() and path.is_dir():
        parser.error(tr(T.cli_output_is_directory, path=path))
    if path.suffix.lower() != ".docx":
        path = Path(str(path) + ".docx")
    return path


def cli_main(argv=None):
    """Handle CLI operations without constructing QApplication or a GUI page."""
    global _FORCE_ENGLISH_UI
    _FORCE_ENGLISH_UI = True
    parser = _build_cli_parser()
    args = parser.parse_args(argv)

    if not HAS_DOCX:
        parser.error(tr(T.no_docx))

    try:
        migrate_legacy_data()
        load_settings()
    except Exception as exc:
        parser.error(tr(T.cli_read_error, err=exc))

    if args.wrapper is not None:
        if not args.wrapper.strip():
            parser.error(tr(T.err_wrapper_empty))
        SETTINGS["placeholder_wrapper"] = args.wrapper

    input_path = Path(args.input_path).expanduser()
    if not input_path.exists():
        parser.error(tr(T.err_template_not_found, path=input_path))
    if not input_path.is_file():
        parser.error(tr(T.cli_input_not_file, path=input_path))
    if input_path.suffix.lower() != ".docx":
        parser.error(tr(T.cli_input_not_docx, path=input_path))
    input_path = input_path.resolve()

    wrapper = _get_wrapper()
    if not isinstance(wrapper, str) or not wrapper:
        parser.error(tr(T.err_wrapper_empty))

    try:
        properties = extract_placeholders(str(input_path))
    except WrapperCountError as exc:
        parser.error(str(exc))
    except Exception as exc:
        parser.error(tr(T.cli_read_error, err=exc))
    variable_names = list(properties.keys())

    if args.count:
        if (args.output_path is not None or args.values is not None
                or args.excel_path is not None or args.header_row or args.force):
            parser.error(tr(T.cli_count_mode_conflict))
        print(tr(T.cli_variable_count, count=len(variable_names)))
        return 0

    if args.header_row and args.excel_path is None:
        parser.error(tr(T.cli_header_only_excel))

    if args.excel_path is not None:
        if args.values is not None or args.force:
            parser.error(tr(T.cli_excel_mode_conflict))
        if not variable_names:
            parser.error(tr(T.cli_no_variables, w=wrapper))
        if not HAS_EXCEL:
            parser.error(tr(T.batch_excel_library_missing))

        try:
            records = read_excel_records(
                args.excel_path, variable_names, header_row=args.header_row
            )
        except ExcelImportError as exc:
            parser.error(str(exc))

        if args.output_path:
            batch_directory = Path(args.output_path).expanduser()
            if batch_directory.suffix.lower() == ".docx":
                parser.error(tr(
                    T.cli_excel_output_folder_error, path=batch_directory
                ))
            if batch_directory.exists() and not batch_directory.is_dir():
                parser.error(tr(
                    T.cli_excel_output_folder_error, path=batch_directory
                ))
        else:
            configured_output = str(SETTINGS.get("output_dir") or "").strip()
            output_root = Path(
                configured_output or (Path.home() / "Desktop")
            ).expanduser()
            batch_directory = output_root / tr(T.batch_output_folder)

        try:
            batch_directory.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            parser.error(tr(T.batch_output_folder_error, err=exc))

        reserved_names = set()
        for row_number, output_name, replacements in records:
            output_file = _safe_batch_filename(
                output_name, batch_directory, reserved_names
            )
            try:
                _save_docx_atomically(
                    input_path,
                    batch_directory / output_file,
                    replacements,
                    {},
                    overwrite=False,
                )
            except Exception as exc:
                parser.error(tr(
                    T.batch_row_write_error, row=row_number, err=exc
                ))

        print(tr(T.batch_success, count=len(records), path=batch_directory))
        return 0

    if not variable_names:
        parser.error(tr(T.cli_no_variables, w=wrapper))
    if not args.output_path:
        parser.error(tr(T.cli_output_required))

    values = args.values or []
    if len(values) > len(variable_names):
        parser.error(tr(
            T.cli_too_many_values,
            expected=len(variable_names),
            actual=len(values),
        ))
    if len(values) < len(variable_names):
        parser.error(tr(
            T.cli_too_few_values,
            expected=len(variable_names),
            actual=len(values),
        ))
    if any(value == "" for value in values):
        parser.error(tr(T.cli_empty_value))

    replacements = dict(zip(variable_names, values))
    output_path = _cli_output_path(args.output_path, parser).resolve()
    if output_path == input_path:
        parser.error(tr(T.cli_output_is_input))
    if output_path.exists() and not args.force:
        parser.error(tr(T.cli_output_exists, path=output_path))

    try:
        _save_docx_atomically(
            input_path, output_path, replacements, {}, overwrite=args.force
        )
    except FileExistsError:
        parser.error(tr(T.cli_output_exists, path=output_path))
    except Exception as exc:
        parser.error(tr(T.cli_write_error, err=exc))

    print(tr(T.cli_success, path=output_path, count=len(variable_names)))
    return 0


def _attach_cli_console_if_needed():
    """Attach CLI streams to a parent console if this launcher lacks them.

    The recommended Windows spec is console-enabled with PyInstaller's
    hide_console='hide-late': this lets CMD wait for CLI completion while an
    owned console is hidden for GUI startup. This helper is only a fallback
    for launchers that start the CLI without standard streams; attaching a
    windowed executable cannot make its parent CMD wait synchronously.
    """
    if sys.platform != "win32" or (sys.stdout is not None and sys.stderr is not None):
        return
    try:
        import ctypes

        kernel32 = ctypes.windll.kernel32
        kernel32.AttachConsole(-1)  # ATTACH_PARENT_PROCESS
        if not kernel32.GetConsoleWindow():
            kernel32.AllocConsole()
        if sys.stdin is None:
            sys.stdin = open("CONIN$", "r", encoding="utf-8", errors="replace")
        if sys.stdout is None:
            sys.stdout = open(
                "CONOUT$", "w", encoding="utf-8", errors="replace", buffering=1
            )
        if sys.stderr is None:
            sys.stderr = open(
                "CONOUT$", "w", encoding="utf-8", errors="replace", buffering=1
            )
    except Exception:
        pass


# ═════════════════════════════════════════════════════════════════
#  GUI startup
# ═════════════════════════════════════════════════════════════════

def gui_main():
    migrate_legacy_data()

    first_run = not SETTINGS_JSON.exists()

    load_settings()

    if sys.platform == "win32":
        myappid = "com.wordcloner.letterissuance.1.0"
        try:
            from ctypes import windll
            windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception as e:
            print(f"[AppUserModelID] failed: {e}")

    app = QApplication([sys.argv[0]])
    icon_path = APP_ROOT / "app.ico"
    app.setWindowIcon(QIcon(str(icon_path)))
    app.setQuitOnLastWindowClosed(False)

    for fam in ("Vazirmatn", "IRANSans", "Tahoma", "Segoe UI"):
        f = QFont(fam, 10)
        if f.exactMatch() or fam == "Segoe UI":
            app.setFont(f)
            break

    if first_run:
        dlg = LanguageSelectDialog()
        if dlg.exec() == QDialog.Accepted and dlg.selected_language:
            SETTINGS["language"] = dlg.selected_language
        save_settings()

    app.setLayoutDirection(_dir())

    current = [None]

    def restart():
        old = current[0]
        try:
            old._save_window_state()
            if hasattr(old, "_window_state_timer"):
                old._window_state_timer.stop()
        except Exception:
            pass
        was_max = old.isMaximized()
        geo = old.normalGeometry() if was_max else old.geometry()
        if not geo.isValid():
            geo = old.geometry()
        old.hide()
        old.deleteLater()

        load_settings()
        app.setLayoutDirection(_dir())

        new_win = MainWindow()
        new_win.theme_changed.connect(restart)
        new_win.setGeometry(geo)
        if was_max:
            new_win.showMaximized()
        else:
            new_win.show()
        current[0] = new_win

    def create_window():
        win = MainWindow()
        win.theme_changed.connect(restart)
        if SETTINGS.get("window_maximized") is True:
            win.showMaximized()
        else:
            win.show()
        return win

    current[0] = create_window()
    return app.exec()


def main(argv=None):
    """Dispatch to CLI when arguments are supplied; otherwise launch the GUI.

    This same entry point is used by the Python console command and the frozen
    WordCloner.exe, so opening the EXE without arguments starts the GUI.
    """
    arguments = list(sys.argv[1:] if argv is None else argv)
    if arguments == ["--gui"]:
        return gui_main()
    if arguments:
        _attach_cli_console_if_needed()
        return cli_main(arguments)
    return gui_main()


if __name__ == "__main__":
    raise SystemExit(main())

'''

## wordcloner_icons.py

'''
"""Downloaded SVG icon loader for WordCloner.

Icon artwork is stored locally in ``resources/icons`` (downloaded from the
Lucide project) and rendered as SVG. The application therefore does not need
network access at runtime.
"""
from __future__ import annotations

import sys
from functools import lru_cache
from pathlib import Path

from PySide6.QtCore import QByteArray, QRectF, QSize, Qt
from PySide6.QtGui import QColor, QIcon, QPainter, QPixmap
from PySide6.QtSvg import QSvgRenderer


C_DARK_GREEN = "#2E5A4C"
C_GREEN = "#3A7A5E"
C_MUTED = "#7A8F86"
C_RED = "#E0453A"


def _application_root() -> Path:
    """Return the directory beside the .py file, or beside a frozen .exe."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


ICON_DIR = _application_root() / "resources" / "icons"
if not ICON_DIR.is_dir():
    # In a regular wheel install, assets are installed in the private package
    # ``wordcloner_assets`` rather than a generic top-level ``resources``.
    try:
        from importlib.resources import files

        packaged_icons = files("wordcloner_assets").joinpath("icons")
        candidate = Path(str(packaged_icons))
        if candidate.is_dir():
            ICON_DIR = candidate
    except (ImportError, ModuleNotFoundError, TypeError, OSError):
        pass


def _color_name(color: str | QColor | None, fallback: str) -> str:
    qcolor = QColor(color if color is not None else fallback)
    return qcolor.name() if qcolor.isValid() else fallback


@lru_cache(maxsize=256)
def _render_svg_icon(name: str, size: int, color: str) -> QIcon:
    """Render one downloaded SVG at the requested size and color."""
    if size <= 0:
        size = 1

    svg_path = ICON_DIR / f"{name}.svg"
    try:
        svg_text = svg_path.read_text(encoding="utf-8")
    except OSError as exc:
        print(f"[icons] Cannot read {svg_path}: {exc}")
        return QIcon()

    # Lucide uses currentColor. Resolve it before passing the SVG to Qt's
    # renderer so rendering is consistent across Qt platforms.
    svg_text = svg_text.replace("currentColor", color)
    renderer = QSvgRenderer(QByteArray(svg_text.encode("utf-8")))
    if not renderer.isValid():
        print(f"[icons] Invalid SVG: {svg_path}")
        return QIcon()

    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing, True)
    painter.setRenderHint(QPainter.SmoothPixmapTransform, True)
    renderer.render(painter, QRectF(0, 0, size, size))
    painter.end()
    return QIcon(pixmap)


def _asset(name: str, size: int, color: str | QColor | None = None,
           fallback: str = C_DARK_GREEN) -> QIcon:
    return _render_svg_icon(name, int(size), _color_name(color, fallback))


# Keep the old public function names/signatures so the UI code can continue
# calling them while the icon artwork comes from local SVG files.
def icon_document(size=22, color=None):
    return _asset("file-text", size, color, C_DARK_GREEN)


def icon_document_big(size=88, color=None):
    # Lucide's file-plus replaces the former hand-drawn page-with-plus glyph.
    return _asset("file-plus", size, color, C_DARK_GREEN)


def icon_plus_circle(size=28, fg="#FFFFFF", bg=None):
    # The downloaded asset is a single-color outlined circle-plus icon.
    # ``bg`` remains accepted for compatibility with the former signature.
    return _asset("circle-plus", size, bg or C_GREEN, C_GREEN)


def icon_plus_white(size=14, color="#FFFFFF"):
    return _asset("plus", size, color, "#FFFFFF")


def icon_trash(size=18, color=None):
    return _asset("trash", size, color, C_MUTED)


def icon_chevron_down(size=16, color=None):
    return _asset("chevron-down", size, color, C_DARK_GREEN)


def icon_grip(size=16, color=None):
    return _asset("grip-horizontal", size, color, C_MUTED)


def icon_upload(size=56, color=None):
    return _asset("upload", size, color, C_GREEN)


def icon_arrow_left(size=18, color=None):
    return _asset("arrow-left", size, color, C_DARK_GREEN)


def icon_gear(size=22, color=None):
    return _asset("settings", size, color, C_DARK_GREEN)


def icon_help(size=20, color=None):
    return _asset("circle-question-mark", size, color, C_GREEN)


def icon_status(size=42, kind="info"):
    status_assets = {
        "info": ("info", C_GREEN),
        "warning": ("triangle-alert", C_GREEN),
        "error": ("circle-x", C_RED),
        "delete": ("trash", C_RED),
        "question": ("circle-question-mark", C_GREEN),
    }
    name, color = status_assets.get(kind, ("info", C_GREEN))
    return _asset(name, size, color, color)


def icon_minimize(size=16, color="#FFFFFF"):
    return _asset("minus", size, color, "#FFFFFF")


def icon_maximize(size=16, color="#FFFFFF"):
    return _asset("square", size, color, "#FFFFFF")


def icon_close(size=14, color="#FFFFFF"):
    return _asset("x", size, color, "#FFFFFF")


def icon_check(size=56, color=None):
    return _asset("check", size, color, C_GREEN)
'''

## wordcloner_paths.py

'''
"""Project-root storage and safe migration from earlier WordCloner folders.

Current data is stored directly beside the application:
    settings.json
    templates.json
    Words/

Earlier locations are copied into the project root without overwriting files.
The originals are retained until the user removes them manually.
"""
from __future__ import annotations

import json
import os
import shutil
import sys
from pathlib import Path


def _application_root() -> Path:
    """Use the app directory, unless the user explicitly selects a data root.

    ``WORDCLONER_HOME`` is useful for a regular/system package installation
    whose site-packages directory is read-only. Without it, source/editable
    installs continue to keep settings and Words beside the application files.
    """
    override = os.environ.get("WORDCLONER_HOME")
    if override and override.strip():
        return Path(override).expanduser().resolve()
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_ROOT = _application_root()
# Kept for compatibility with the existing file-picker start-directory code.
DOCUMENTS_DIR = APP_ROOT

# Data now lives directly at the project root.
WORDCLONER_DIR = APP_ROOT
WORDS_DIR = APP_ROOT / "Words"
TEMPLATES_JSON = APP_ROOT / "templates.json"
SETTINGS_JSON = APP_ROOT / "settings.json"

# Both locations may contain data from earlier versions.
PREVIOUS_WORDCLONER_DIR = APP_ROOT / "WordCloner"
LEGACY_WORDCLONER_DIR = Path.home() / "Documents" / "WordCloner"


def _copy_missing_tree(source: Path, destination: Path) -> bool:
    """Copy missing files recursively; never overwrite an existing target."""
    copied_any = False
    if not source.is_dir():
        return False

    destination.mkdir(parents=True, exist_ok=True)
    try:
        children = list(source.iterdir())
    except OSError as exc:
        print(f"[migration] Cannot read {source}: {exc}")
        return False

    for child in children:
        target = destination / child.name
        try:
            if child.is_dir():
                if target.exists() and not target.is_dir():
                    print(f"[migration] Kept existing file; skipped folder: {target}")
                    continue
                copied_any = _copy_missing_tree(child, target) or copied_any
            elif not target.exists():
                shutil.copy2(str(child), str(target))
                copied_any = True
        except Exception as exc:
            print(f"[migration] Could not copy {child} to {target}: {exc}")
    return copied_any


def _relocated_address(address):
    """Map old absolute template paths to the new root-level location."""
    if not isinstance(address, str) or not address.strip():
        return address

    try:
        old_path = Path(address).expanduser().resolve()
    except (OSError, RuntimeError):
        return address

    roots = (PREVIOUS_WORDCLONER_DIR, LEGACY_WORDCLONER_DIR)
    for old_root in roots:
        try:
            old_root_resolved = old_root.resolve()
            relative = old_path.relative_to(old_root_resolved)
        except (OSError, ValueError, RuntimeError):
            continue
        return str(APP_ROOT / relative)

    return address


def _rewrite_template_addresses() -> bool:
    """Merge template records from both legacy folders and relocate paths."""
    templates = []
    changed = False

    if TEMPLATES_JSON.is_file():
        try:
            with TEMPLATES_JSON.open("r", encoding="utf-8") as stream:
                current = json.load(stream)
            if isinstance(current, list):
                templates = [item for item in current if isinstance(item, dict)]
            else:
                print(f"[migration] Invalid template list in {TEMPLATES_JSON}")
                return False
        except Exception as exc:
            print(f"[migration] Could not read {TEMPLATES_JSON}: {exc}")
            return False

    def identity(template):
        name = (template.get("name") or "").strip()
        if name:
            return ("name", name.casefold())
        address = template.get("address")
        return ("address", str(address or "").casefold())

    identities = set()
    for template in templates:
        old_address = template.get("address")
        new_address = _relocated_address(old_address)
        if new_address != old_address:
            template["address"] = new_address
            changed = True
        identities.add(identity(template))

    # Merge any templates not already present. Existing root-level records win.
    for source_root in (PREVIOUS_WORDCLONER_DIR, LEGACY_WORDCLONER_DIR):
        source_file = source_root / "templates.json"
        try:
            if not source_file.is_file() or source_file.resolve() == TEMPLATES_JSON.resolve():
                continue
            with source_file.open("r", encoding="utf-8") as stream:
                source_templates = json.load(stream)
            if not isinstance(source_templates, list):
                continue
            for source_template in source_templates:
                if not isinstance(source_template, dict):
                    continue
                template = dict(source_template)
                old_address = template.get("address")
                new_address = _relocated_address(old_address)
                if new_address != old_address:
                    template["address"] = new_address
                key = identity(template)
                if key in identities:
                    continue
                templates.append(template)
                identities.add(key)
                changed = True
        except Exception as exc:
            print(f"[migration] Could not merge templates from {source_file}: {exc}")

    if changed:
        try:
            with TEMPLATES_JSON.open("w", encoding="utf-8") as stream:
                json.dump(templates, stream, ensure_ascii=False, indent=4)
        except Exception as exc:
            print(f"[migration] Could not write {TEMPLATES_JSON}: {exc}")
            return False
    return changed


def migrate_legacy_data() -> bool:
    """Bring old WordCloner contents directly into the project root.

    Migration sources are processed in order: the prior project-local
    ``WordCloner/`` folder first, then ``Documents/WordCloner``. Existing
    root-level files always win; missing files are copied. Source folders are
    deliberately retained as backups. Old absolute addresses in
    ``templates.json`` are then rewritten to ``<project-root>/Words/...``.
    """
    migrated = False
    sources = (PREVIOUS_WORDCLONER_DIR, LEGACY_WORDCLONER_DIR)
    root_resolved = APP_ROOT.resolve()

    for source in sources:
        try:
            if not source.is_dir() or source.resolve() == root_resolved:
                continue
            if _copy_missing_tree(source, APP_ROOT):
                migrated = True
                print(f"[migration] Copied missing data from: {source}")
        except Exception as exc:
            print(f"[migration] Could not migrate {source}: {exc}")

    migrated = _rewrite_template_addresses() or migrated
    if migrated:
        print(f"[migration] Active data location: {APP_ROOT}")
        print("[migration] Original WordCloner folders were left untouched.")
    return migrated
'''

## wordcloner_translations.py

'''
"""Persian and English UI copy for WordCloner.

Wording is standardized for an administrative/document-issuance workflow.
"""


def _entry(persian, english):
    return {"persian": persian, "english": english}


class T:
    err_no_vars= _entry(
        (
        "در این قالب هیچ متغیری پیدا نشد.\n"
        "مطمئن شوید متغیرها با {w} احاطه شده‌اند "
        "(مثلاً {w}نام مشتری{w})."
    ),
    (
        "No variables were found in this template.\n"
        "Make sure variables are wrapped with {w} "
        "(for example {w}customer_name{w})."
    )
    )
    # Application and navigation
    app_title = _entry("سامانه صدور نامه", "Letter Issuance System")
    settings = _entry("تنظیمات", "Settings")
    settings_spaced = _entry("  تنظیمات", "  Settings")
    issue_letter = _entry("صدور نامه", "Issue Letter")
    add_new_letter = _entry("ثبت قالب نامهٔ جدید", "Register New Letter Template")
    back = _entry("بازگشت", "Back")
    back_spaced = _entry("  بازگشت", "  Back")

    # Template selection and quick access
    select_template = _entry("انتخاب قالب نامه", "Select a Letter Template")
    add_letter = _entry("ثبت قالب نامه", "Register Template")
    search_template = _entry("جست‌وجوی قالب نامه...", "Search letter templates...")
    untitled = _entry("بدون عنوان", "Untitled")
    drag_to_reorder = _entry(
        "برای تغییر ترتیب، بکشید و رها کنید.",
        "Drag and drop to change the order.",
    )

    quick_access = _entry("دسترسی سریع", "Quick Access")
    add_to_qa = _entry("مدیریت دسترسی سریع", "Manage Quick Access")
    add_template_to_qa = _entry(
        "افزودن قالب به دسترسی سریع", "Add Template to Quick Access"
    )
    remove_from_qa = _entry("حذف از دسترسی سریع", "Remove from Quick Access")
    no_qa_templates = _entry(
        "قالبی به دسترسی سریع افزوده نشده است؛ برای افزودن، «+» را انتخاب کنید.",
        "No templates have been added to Quick Access. Select “+” to add one.",
    )
    select_qa_hint = _entry(
        "قالب‌هایی را که باید در بخش «دسترسی سریع» نمایش داده شوند انتخاب کنید:",
        "Select the templates to display in Quick Access:",
    )
    no_templates_yet = _entry(
        "هنوز قالبی ثبت نشده است.", "No templates have been registered yet."
    )
    save = _entry("ذخیره", "Save")
    cancel = _entry("انصراف", "Cancel")

    # Word file selection
    drop_zone_hint = _entry(
        "فایل Word را اینجا رها کنید\nیا برای انتخاب از رایانه کلیک کنید",
        "Drop a Word file here\nor click to browse your computer",
    )
    select_word_file = _entry("انتخاب فایل Word", "Select a Word File")
    word_file_filter = _entry("سند Word (*.docx)", "Word Document (*.docx)")

    # Template fields and formatting
    font_size = _entry("اندازهٔ قلم", "Font Size")
    auto = _entry("خودکار", "Auto")
    bold = _entry("پررنگ", "Bold")
    italic = _entry("مورب", "Italic")
    color = _entry("رنگ متن", "Text Color")
    color_tooltip = _entry(
        "انتخاب رنگ متن؛ رنگ پیش‌فرض مشکی است.",
        "Choose a text color; the default is black.",
    )
    pick_color = _entry("انتخاب رنگ متن", "Select Text Color")
    font_size_tooltip = _entry(
        "اندازهٔ قلم را مشخص کنید. گزینهٔ «خودکار» اندازهٔ متن مجاور را اعمال می‌کند.",
        "Choose a font size. “Auto” matches the surrounding text.",
    )
    letter_name = _entry("نام قالب نامه", "Template Name")
    word_template_file = _entry("فایل Word قالب نامه", "Word Template File")
    output_filename = _entry("نام فایل خروجی", "Output File Name")
    variable_values = _entry("مقادیر متغیرها", "Variable Values")
    confirm_save = _entry("تأیید و ثبت قالب", "Confirm and Register Template")
    confirm_issue = _entry("تأیید و صدور نامه", "Confirm and Issue Letter")
    batch_excel_issue = _entry(
        "صدور جمعی از اکسل", "Bulk Issue from Excel"
    )
    batch_issue_title = _entry(
        "صدور جمعی از اکسل", "Bulk Issue from Excel"
    )
    batch_explanation = _entry(
        "در شیت فعال اکسل، ستون اول نام فایل Word خروجی است. "
        "ستون دوم به متغیر اول، ستون سوم به متغیر دوم و به همین ترتیب "
        "به همهٔ متغیرهای قالب اختصاص دارد. هر ردیف یک سند می‌سازد.\n"
        "سلول خالی برای متغیر خالی پذیرفته می‌شود. می‌توانید سطر اول را "
        "عنوان ستون‌ها قرار دهید و گزینهٔ مربوط را فعال کنید.",
        "In the active Excel sheet, column 1 is the output Word filename. "
        "Column 2 maps to the first template variable, column 3 to the second, "
        "and so on. Each row creates one document.\n"
        "An empty cell leaves that variable blank. You may use row 1 for column "
        "headings and enable the header-row option.",
    )
    batch_columns_title = _entry(
        "ترتیب ستون‌ها در اکسل", "Excel column mapping"
    )
    batch_format_title = _entry(
        "قالب‌بندی هر متغیر در سندهای خروجی", "Formatting for each variable"
    )
    batch_mapping_filename = _entry(
        "ستون {column}: نام فایل خروجی",
        "Column {column}: output filename",
    )
    batch_mapping_variable = _entry(
        "ستون {column}: متغیر «{name}»",
        "Column {column}: variable “{name}”",
    )
    batch_header_row = _entry(
        "سطر اول عنوان ستون‌هاست (در پردازش نادیده گرفته شود)",
        "Row 1 contains headings (skip it)",
    )
    batch_drop_hint = _entry(
        "فایل اکسل را اینجا رها کنید\nیا برای انتخاب از رایانه کلیک کنید",
        "Drop an Excel workbook here\nor click to browse your computer",
    )
    batch_excel_filter = _entry(
        "فایل اکسل (*.xlsx *.xlsm)", "Excel workbook (*.xlsx *.xlsm)"
    )
    batch_confirm = _entry(
        "تأیید و صدور جمعی", "Confirm and Create Documents"
    )
    batch_output_folder = _entry("صدور جمعی", "Bulk Issuance")
    batch_excel_missing = _entry(
        "ابتدا فایل اکسل را انتخاب یا در کادر مشخص‌شده رها کنید.",
        "Select or drop an Excel workbook first.",
    )
    batch_excel_library_missing = _entry(
        "کتابخانهٔ openpyxl نصب نشده است. برای نصب، دستور pip install openpyxl را اجرا کنید.",
        "The openpyxl library is missing. Install it with: pip install openpyxl",
    )
    batch_excel_not_found = _entry(
        "فایل اکسل یافت نشد:\n{path}",
        "Excel workbook not found:\n{path}",
    )
    batch_excel_type_error = _entry(
        "فایل باید با قالب .xlsx یا .xlsm باشد:\n{path}",
        "The workbook must be an .xlsx or .xlsm file:\n{path}",
    )
    batch_excel_read_error = _entry(
        "خواندن فایل اکسل ناموفق بود:\n{err}",
        "Could not read the Excel workbook:\n{err}",
    )
    batch_no_data_rows = _entry(
        "در شیت فعال، ردیف داده‌ای برای صدور پیدا نشد.",
        "No data rows were found in the active worksheet.",
    )
    batch_filename_empty = _entry(
        "نام فایل خروجی در سطر {row} خالی است.",
        "The output filename is empty on row {row}.",
    )
    batch_extra_columns = _entry(
        "سطر {row} در ستون {column} دادهٔ اضافه دارد. برای {variables} متغیر، "
        "پس از ستون نام فایل فقط {variables} ستون متغیر لازم است.",
        "Row {row} has extra data in column {column}. With {variables} variable(s), "
        "only {variables} variable column(s) may follow the filename column.",
    )
    batch_output_folder_error = _entry(
        "پوشهٔ خروجی صدور جمعی ایجاد یا قابل‌استفاده نیست:\n{err}",
        "Could not create or access the bulk output folder:\n{err}",
    )
    batch_row_write_error = _entry(
        "صدور سند در سطر {row} ناموفق بود:\n{err}",
        "Could not issue the document on row {row}:\n{err}",
    )
    batch_success = _entry(
        "صدور جمعی انجام شد. {count} سند در پوشهٔ زیر ذخیره شد:\n{path}",
        "Bulk issuance completed. {count} document(s) were saved in:\n{path}",
    )
    ph_letter_name = _entry(
        "برای مثال: نامهٔ پذیرش مقاله", "e.g., Article Acceptance Letter"
    )
    ph_filename = _entry(
        "برای مثال: نامه_پذیرش_مقاله", "e.g., Article_Acceptance_Letter"
    )
    no_variables = _entry(
        "برای این قالب متغیری تعریف نشده است.",
        "No variables are defined for this template.",
    )
    issue_named = _entry("صدور نامه: {name}", "Issue Letter: {name}")

    # Settings
    app_language = _entry("زبان رابط کاربری", "Interface Language")
    guide = _entry("  راهنما", "  Guide")
    guide_tooltip = _entry(
        "اجرای راهنمای ویدیویی WordCloner.",
        "Open the WordCloner video guide.",
    )
    output_path = _entry(
        "محل ذخیرهٔ نامه‌های صادرشده", "Output Folder for Issued Letters"
    )
    placeholder_wrapper = _entry("نشانگر متغیرها", "Placeholder Wrapper")
    browse = _entry("انتخاب...", "Browse...")
    save_settings = _entry("ذخیرهٔ تنظیمات", "Save Settings")
    select_output_dir = _entry("انتخاب پوشهٔ خروجی", "Select Output Folder")
    ph_wrapper = _entry("برای مثال: \"\"", 'e.g., ""')
    wrapper_hint = _entry(
        "متغیرها در فایل Word میان این نشانگر قرار می‌گیرند.\nنمونه: {example}",
        "In the Word file, place each variable between these wrappers.\nExample: {example}",
    )
    user_name = _entry("نام کاربر", "User Name")

    # Dialog buttons and messages
    ok = _entry("تأیید", "OK")
    yes = _entry("بله", "Yes")
    delete = _entry("حذف", "Delete")
    missing_library = _entry("کتابخانهٔ موردنیاز یافت نشد", "Required Library Not Found")
    no_docx = _entry(
        "کتابخانهٔ python-docx نصب نشده است.\n\nبرای نصب، دستور زیر را اجرا کنید:\npip install python-docx",
        "The python-docx library is not installed.\n\nInstall it by running:\npip install python-docx",
    )
    error = _entry("خطا", "Error")
    err_empty_name = _entry(
        "لطفاً نام قالب نامه را وارد کنید.", "Enter a name for the letter template."
    )
    err_empty_file = _entry(
        "لطفاً فایل Word قالب را انتخاب کنید یا آن را در ناحیهٔ مشخص‌شده رها کنید.",
        "Select a Word template file or drop it in the designated area.",
    )
    err_name_exists = _entry(
        "قالبی با نام «{name}» از پیش ثبت شده است.",
        "A template named “{name}” is already registered.",
    )
    saved = _entry("ثبت قالب انجام شد", "Template Registered")
    saved_with_n = _entry(
        "قالب نامهٔ «{name}» با {n} متغیر ثبت شد.",
        "Template “{name}” was registered with {n} variable(s).",
    )
    saved_no_vars = _entry(
        "یادآوری: در سند انتخاب‌شده، متغیری با نشانگر «{w}» یافت نشد.",
        "Note: No placeholders wrapped with “{w}” were found in the selected document.",
    )
    err_wrapper_count = _entry(
        "تعداد نشانگرهای متغیر با هم مطابقت ندارد",
        "Unmatched Placeholder Wrappers",
    )
    err_generic = _entry("خطایی رخ داد:\n{err}", "An error occurred:\n{err}")
    err_empty_filename = _entry(
        "لطفاً نام فایل خروجی را وارد کنید.", "Enter a name for the output file."
    )
    err_fill_one = _entry(
        "لطفاً مقدار دست‌کم یک متغیر را وارد کنید.",
        "Enter a value for at least one variable.",
    )
    err_output_dir_create = _entry(
        "ایجاد پوشهٔ خروجی انجام نشد:\n{err}",
        "Could not create the output folder:\n{err}",
    )
    err_template_not_found = _entry(
        "فایل قالب یافت نشد:\n{path}", "Template file not found:\n{path}"
    )
    file_exists = _entry("فایل از قبل وجود دارد", "File Already Exists")
    file_overwrite = _entry(
        "فایل «{f}» در پوشهٔ خروجی از قبل وجود دارد.\nآیا می‌خواهید آن را جایگزین کنید؟",
        "The file “{f}” already exists in the output folder.\nDo you want to replace it?",
    )
    issued = _entry("صدور نامه انجام شد", "Letter Issued")
    issued_ok = _entry(
        "نامه با موفقیت در مسیر زیر ذخیره شد:\n{path}",
        "The letter was issued successfully and saved to:\n{path}",
    )
    err_issue = _entry(
        "صدور نامه با خطا مواجه شد:\n{err}",
        "The letter could not be issued:\n{err}",
    )
    delete_template = _entry("حذف قالب نامه", "Delete Letter Template")
    confirm_delete_tpl = _entry(
        "آیا از حذف قالب نامهٔ «{name}» اطمینان دارید؟",
        "Are you sure you want to delete the template “{name}”?",
    )
    deleted = _entry("حذف قالب انجام شد", "Template Deleted")
    deleted_ok = _entry(
        "قالب نامهٔ «{name}» با موفقیت حذف شد.",
        "Template “{name}” was deleted successfully.",
    )
    err_output_empty = _entry(
        "لطفاً پوشهٔ خروجی را مشخص کنید.", "Select an output folder."
    )
    err_output_invalid = _entry(
        "پوشهٔ خروجی قابل استفاده نیست:\n{err}",
        "The output folder is not accessible:\n{err}",
    )
    err_wrapper_empty = _entry(
        "نشانگر متغیرها نمی‌تواند خالی باشد.",
        "The placeholder wrapper cannot be empty.",
    )
    var_count_error_detail = _entry(
        "تعداد نشانگرهای «{w}» در سند زوج نیست (تعداد: {n}).\n"
        "ممکن است نشانگر باز یا بستهٔ متناظر آن درج نشده باشد.\n"
        "لطفاً هر متغیر را به این صورت وارد کنید: {w}نام_متغیر{w}.",
        "The number of placeholder wrappers “{w}” in the document must be even "
        "(found: {n}).\nAn opening or closing wrapper may be missing.\n"
        "Write each variable as {w}variable_name{w}.",
    )

    # Command-line document issuance
    cli_values_required = _entry(
        "مقدارهای متغیرها را با گزینهٔ -v وارد کنید.",
        "Provide the variable values with -v.",
    )
    cli_too_many_values = _entry(
        "تعداد مقدارهای واردشده ({actual}) از تعداد متغیرهای قالب ({expected}) بیشتر است.",
        "Too many values: supplied {actual}, but the template has {expected} unique variable(s).",
    )
    cli_too_few_values = _entry(
        "تعداد مقدارهای واردشده ({actual}) از تعداد متغیرهای قالب ({expected}) کمتر است.\n"
        "برای متغیرِ خالی، جایگاه آن را حذف نکنید؛ به‌جای آن دقیقاً یک فاصله وارد کنید: \" \".",
        "Too few values: supplied {actual}, but the template has {expected} unique variable(s).\n"
        "Do not omit a blank variable's position; pass one space instead: \" \".",
    )
    cli_empty_value = _entry(
        "برای مقدار خالی از یک فاصلهٔ داخل نقل‌قول استفاده کنید (\" \")، نه رشتهٔ خالی (\"\").",
        "Use a quoted single space (\" \"), not an empty string (\"\"), for a blank value.",
    )
    cli_variable_count = _entry(
        "تعداد متغیرهای یکتای قالب: {count}",
        "Unique variables in template: {count}",
    )
    cli_count_mode_conflict = _entry(
        "گزینهٔ --count فقط با -i و -w قابل استفاده است؛ آن را با گزینه‌های صدور ترکیب نکنید.",
        "--count accepts -i and optional -w only; do not combine it with issuance options.",
    )
    cli_excel_mode_conflict = _entry(
        "برای صدور از اکسل، گزینه‌های -v و --force را وارد نکنید؛ مقدارها از فایل اکسل خوانده می‌شوند و نام‌های تکراری خودکار یکتا می‌شوند.",
        "Do not combine --excel with -v or --force; values come from the workbook and duplicate filenames are made unique.",
    )
    cli_output_required = _entry(
        "برای صدور تکی، مسیر فایل خروجی را با گزینهٔ -o مشخص کنید.",
        "For a single document, specify the output file with -o.",
    )
    cli_header_only_excel = _entry(
        "گزینهٔ --header-row فقط همراه با --excel قابل استفاده است.",
        "--header-row can only be used with --excel.",
    )
    cli_excel_output_folder_error = _entry(
        "در حالت اکسل، گزینهٔ -o باید مسیر یک پوشه باشد، نه فایل .docx:\n{path}",
        "In Excel mode, -o must be a folder path, not a .docx file:\n{path}",
    )
    cli_no_variables = _entry(
        "در فایل قالب هیچ متغیری پیدا نشد. مطمئن شوید متغیرها با نشانگر «{w}» نوشته شده‌اند.",
        "The template contains no variables. Check that each variable is wrapped with “{w}”.",
    )
    cli_value_count = _entry(
        "قالب {expected} متغیر یکتا دارد، اما {actual} مقدار وارد شده است.\n"
        "مقادیر را به ترتیب نخستین ظهور متغیرها در سند وارد کنید.",
        "The template has {expected} unique variable(s), but {actual} value(s) were supplied.\n"
        "Enter values in the order the variables first appear in the document.",
    )
    cli_input_not_file = _entry(
        "مسیر ورودی یک فایل نیست:\n{path}",
        "The input path is not a file:\n{path}",
    )
    cli_input_not_docx = _entry(
        "فایل ورودی باید با قالب .docx باشد:\n{path}",
        "The input must be a .docx Word document:\n{path}",
    )
    cli_output_is_input = _entry(
        "مسیر فایل خروجی با فایل قالب یکسان است؛ فایل قالب بازنویسی نمی‌شود.",
        "The output path is the same as the template; the template will not be overwritten.",
    )
    cli_output_is_directory = _entry(
        "مسیر خروجی به یک پوشه اشاره می‌کند، نه فایل:\n{path}",
        "The output path points to a directory, not a file:\n{path}",
    )
    cli_output_exists = _entry(
        "فایل خروجی از قبل وجود دارد:\n{path}\nبرای جایگزینی آن، گزینهٔ --force را اضافه کنید.",
        "The output file already exists:\n{path}\nUse --force to replace it.",
    )
    cli_output_directory_error = _entry(
        "پوشهٔ خروجی ایجاد یا قابل دسترسی نشد:\n{err}",
        "Could not create or access the output folder:\n{err}",
    )
    cli_read_error = _entry(
        "خواندن فایل قالب ناموفق بود:\n{err}",
        "Could not read the template:\n{err}",
    )
    cli_write_error = _entry(
        "ذخیرهٔ فایل خروجی ناموفق بود:\n{err}",
        "Could not save the output file:\n{err}",
    )
    cli_success = _entry(
        "سند با موفقیت در این مسیر ذخیره شد:\n{path}\nتعداد متغیرها: {count}",
        "Document created successfully:\n{path}\nVariables filled: {count}",
    )
    guide_not_found = _entry(
        "فایل راهنمای ویدیویی یافت نشد:\n{path}",
        "The video guide was not found:\n{path}",
    )
    guide_open_error = _entry(
        "اجرای فایل راهنما ناموفق بود:\n{err}",
        "Could not open the video guide:\n{err}",
    )
'''

# End of source code

# 🤖 AI-Readme.md — WordCloner Knowledge Base for AI Assistants

> **📌 Purpose of this file**
>
> This document is a **comprehensive context primer** for any AI assistant (ChatGPT, Claude, Gemini, DeepSeek, local LLMs, etc.) that is asked a question about the **WordCloner** project.
>
> When a user or developer sends you this file, treat it as the **authoritative reference** for how WordCloner is structured, how it behaves, and how to extend it. Prefer the information here over generic assumptions.
>
> **How to use it**
>
> 1. The user will typically paste **this entire file** along with the `WordCloner.py` source code at the top.
> 2. Read this document first — it maps every module, every data structure, and every entry point.
> 3. When answering, cite the module name and function name whenever possible.
> 4. If the user's question is ambiguous (GUI vs CLI, Persian vs English), **ask for clarification** before answering.
> 5. Never invent features that are not documented here. If something is unknown, say so and suggest reading the actual source.

---

## 📖 Table of Contents

1. [Project Summary](#1--project-summary)
2. [High-Level Architecture](#2--high-level-architecture)
3. [File-by-File Reference](#3--file-by-file-reference)
4. [Data Models (JSON Schemas)](#4--data-models-json-schemas)
5. [The Placeholder Pipeline](#5--the-placeholder-pipeline)
6. [The GUI Layer](#6--the-gui-layer)
7. [The CLI Layer](#7--the-cli-layer)
8. [Internationalization (i18n)](#8--internationalization-i18n)
9. [Settings System](#9--settings-system)
10. [Icon & Resource System](#10--icon--resource-system)
11. [Paths & Legacy Migration](#11--paths--legacy-migration)
12. [Batch Issuance from Excel](#12--batch-issuance-from-excel)
13. [System Tray Integration](#13--system-tray-integration)
14. [Build & Packaging](#14--build--packaging)
15. [Installer (Inno Setup)](#15--installer-inno-setup)
16. [Common User Questions & Answers](#16--common-user-questions--answers)
17. [Common Developer Extension Points](#17--common-developer-extension-points)
18. [Known Limitations & Gotchas](#18--known-limitations--gotchas)
19. [Debugging & Diagnostics](#19--debugging--diagnostics)
20. [Glossary](#20--glossary)
21. [How an AI Should Answer Questions](#21--how-an-ai-should-answer-questions)

---

## 1. 🎯 Project Summary

**WordCloner** is a Windows desktop application that generates personalized Word (`.docx`) documents from templates. It is written in **Python 3.10+** using:

- **PySide6** (Qt 6) for the GUI.
- **python-docx** for reading and writing `.docx` templates.
- **openpyxl** for reading Excel workbooks (batch mode).
- **PyInstaller** for producing a standalone `.exe`.
- **Inno Setup** for the installer.

**Primary design goal:** drastically reduce the time required to issue **repetitive letters and circulars (ابلاغیه)** in administrative offices and departments. The same `.docx` template can be reused hundreds of times with different variable values, either manually (GUI), via Excel (bulk mode), or via the CLI.

**Two execution modes:**
- **GUI mode** — launched when the executable is started **without arguments**.
- **CLI mode** — launched when **any argument** is passed (or explicitly with `--gui` for the GUI).

**Key user-facing concepts:**
- **Template** — a `.docx` file containing placeholder variables wrapped by a marker (default `##`).
- **Variable** — any text wrapped between two markers, e.g. `##customer_name##`.
- **Quick Access** — a pinned shortlist of templates.
- **Batch issuance** — producing many documents from one Excel workbook.

---

## 2. 🏗️ High-Level Architecture

```
                 ┌──────────────────────────────────┐
                 │            main()                 │
                 │  (dispatch GUI vs CLI)            │
                 └────────────┬──────────────┬──────┘
                              │              │
                 (no args)    │              │   (any arg)
                              ▼              ▼
                        gui_main()        cli_main()
                              │              │
                              ▼              ▼
                 ┌────────────────┐   ┌──────────────────┐
                 │  MainWindow    │   │ argparse parser  │
                 │  (QMainWindow) │   │ → extract_placeholders
                 └──────┬─────────┘   │ → replace_placeholders
                        │             │ → _save_docx_atomically
                        ▼             └──────────────────┘
        ┌───────────────┼───────────────┬────────────────┐
        │               │               │                │
   ┌────▼───┐     ┌─────▼─────┐   ┌─────▼─────┐   ┌──────▼──────┐
   │HomePage│     │AddLetter  │   │EditLetter │   │BulkIssue    │
   │        │     │Page       │   │Page       │   │Page         │
   └────────┘     └───────────┘   └───────────┘   └─────────────┘
        │
        ▼
   ┌─────────────────────┐
   │ TemplateDropdown /  │
   │ QuickAccessWidget   │
   └─────────────────────┘
```

**Shared services used by both GUI and CLI:**

- `extract_placeholders(docx_path)` — returns a dict of variables found in a template.
- `replace_placeholders_in_docx(src, dst, replacements, font_overrides)` — writes the output.
- `_save_docx_atomically(...)` — atomic write via a temporary sibling file.
- `read_excel_records(...)` — parse an Excel workbook into per-row replacements.
- `load_templates()` / `save_templates()` — template library persistence.
- `load_settings()` / `save_settings()` — app settings persistence.
- `migrate_legacy_data()` — one-time migration from old locations.
- `tr(...)` — translation lookup helper.

---

## 3. 📂 File-by-File Reference

### `src/WordCloner.py` — main module (approx. 1900+ lines)

Contains **everything** except icons, paths, and translations:

| Section | Purpose |
|---------|---------|
| Imports + optional deps | `HAS_DOCX`, `HAS_EXCEL` flags. |
| `load_settings` / `save_settings` | Settings persistence. |
| `tr(item, **kwargs)` | Translate a `_entry()` dict. |
| `_is_rtl()` / `_dir()` | Decide RTL/LTR layout. |
| `ComboWithArrow` | QComboBox with custom chevron. |
| `ThemedDialog` | Frameless dialog (info/warning/error/confirm/delete). |
| `QuickAccessPicker` | Dialog for selecting quick-access templates. |
| `load_templates_raw` / `load_templates` / `save_templates` | Template JSON I/O. |
| `normalize_quotes` | Normalize fancy quotes to ASCII `"`. |
| `_get_wrapper()` | Read `placeholder_wrapper` from settings. |
| `WrapperCountError` | Raised when markers are unmatched. |
| `compute_auto_font` | Infer font size from surrounding runs. |
| `_nearest_run` | Find the run closest to a placeholder. |
| `_copy_run_fonts` | Copy `w:rFonts` from source run to destination. |
| `_template_paragraphs_in_document_order` | Body + table paragraphs. |
| `extract_placeholders` | **Core extractor.** |
| `replace_placeholders_in_docx` | **Core replacer.** |
| `ExcelImportError` | Domain error for Excel issues. |
| `_excel_is_blank`, `_excel_cell_text` | Cell normalization. |
| `read_excel_records` | Excel → records. |
| `_save_docx_atomically` | Temp-file + `os.replace`. |
| `_safe_batch_filename` | Sanitize + dedupe filenames. |
| `_font_size_value` | Convert combo text to `float` or `None`. |
| `TitleBar` | Custom frameless title bar. |
| `DropdownButton`, `DragGhost`, `GripLabel`, `DragController` | Drag & drop infrastructure. |
| `QuickAccessDragController` | Specialized drag controller. |
| `PopupItem`, `DropdownPopup`, `TemplateDropdown` | Template dropdown. |
| `QuickAccessItem`, `QuickAccessWidget` | Quick Access section. |
| `DropZone`, `ExcelDropZone` | File drop zones. |
| `RoundToggle`, `FontSizeCombo` | Small input widgets. |
| `HomePage`, `AddLetterPage`, `EditLetterPage`, `BulkIssuePage`, `SettingsPage` | The five stacked pages. |
| `Container` | Rounded-rect painted background. |
| `MainWindow` | Assembles everything, tray, resizing. |
| `LangButton`, `LanguageSelectDialog` | First-run language chooser. |
| `_build_cli_parser`, `cli_main` | CLI entry. |
| `_attach_cli_console_if_needed` | Windows console attach. |
| `gui_main` | GUI entry. |
| `main` | Dispatcher. |

### `src/wordcloner_paths.py`

Resolves `APP_ROOT`, `WORDS_DIR`, `TEMPLATES_JSON`, `SETTINGS_JSON`. Handles legacy migration from two previous folders:
- `APP_ROOT / "WordCloner"` (older project-local).
- `~/Documents/WordCloner` (very old location).

**Environment override:** set `WORDCLONER_HOME` to redirect all data to a custom folder (useful for read-only installs).

### `src/wordcloner_icons.py`

Loads SVG icons from `resources/icons/` (or a packaged `wordcloner_assets` module in some distributions). Renders them at requested size and color using `QSvgRenderer`. Caches results with `lru_cache`.

Public functions: `icon_document`, `icon_document_big`, `icon_plus_circle`, `icon_plus_white`, `icon_trash`, `icon_chevron_down`, `icon_grip`, `icon_upload`, `icon_arrow_left`, `icon_gear`, `icon_help`, `icon_status`, `icon_minimize`, `icon_maximize`, `icon_close`, `icon_check`.

### `src/wordcloner_translations.py`

Defines `T` as a class of `_entry(persian, english)` dicts. Usage: `tr(T.some_key, **kwargs)` where `kwargs` fill `{name}` style placeholders.

### `src/WordCloner.spec`

PyInstaller build spec. Recommended settings:
- `console=True` with `hide_console='hide-late'` so CLI works while GUI hides the console.
- Bundle `resources/icons` and `Guide/` via `datas`.
- Include `PySide6`, `python-docx`, `openpyxl`, `numpy`, `Pillow` etc.

### `resources/icons/*.svg`

Lucide-format icons. Only these are bundled: `arrow-left`, `check`, `chevron-down`, `circle-plus`, `circle-question-mark`, `circle-x`, `file-plus`, `file-text`, `grip-horizontal`, `info`, `minus`, `plus`, `settings`, `square`, `trash`, `triangle-alert`, `upload`, `x`.

### `Practice/`

Sample templates and Excel data:
- `WordCloner_Test_Template_EN.docx` — English template.
- `WordCloner_Test_Template_FA.docx` — Persian template.
- `WordCloner_Test_Data_EN.xlsx` — matching batch data.
- `WordCloner_Test_Data_FA.xlsx` — matching batch data.

### `Guide/Persian_guide.mp4`, `Guide/English_guide.mp4`

Video tutorials opened via the **راهنما / Guide** button.

---

## 4. 🗃️ Data Models (JSON Schemas)

### `settings.json`

Located at `APP_ROOT / "settings.json"`.

```json
{
    "output_dir": "C:\\Users\\Ali\\Desktop",
    "placeholder_wrapper": "##",
    "quick_access": ["Article Acceptance", "Invoice"],
    "language": "fa",
    "window_size": [900, 640],
    "window_maximized": false
}
```

| Key | Type | Meaning |
|-----|------|---------|
| `output_dir` | string | Where issued `.docx` files are written. Default: user's Desktop. |
| `placeholder_wrapper` | string | Marker used to detect variables. Default `##`. |
| `quick_access` | string[] | Names of templates shown in Quick Access. |
| `language` | `"fa"` or `"en"` | UI language. |
| `window_size` | `[w, h]` | Last non-maximized size. |
| `window_maximized` | bool | Whether the window was maximized on close. |

**Load precedence:**
1. Start from `DEFAULT_SETTINGS`.
2. If `settings.json` exists and is a dict, `update()` with it.
3. Sanitize: `quick_access` must be a list, `language` must be `"fa"` or `"en"`.

### `templates.json`

A **list** of template records:

```json
[
    {
        "name": "Article Acceptance",
        "address": "C:\\...\\Words\\acceptance.docx",
        "Properties": {
            "author_name": {
                "start": 42,
                "auto_size": 12.0,
                "font_name": "B Nazanin"
            },
            "paper_title": {
                "start": 87,
                "auto_size": null,
                "font_name": null
            }
        }
    }
]
```

| Field | Meaning |
|-------|---------|
| `name` | Display name. Must be unique across the library. |
| `address` | Absolute path to the copied `.docx` inside `Words/`. |
| `Properties` | Map of variable name → metadata. |
| `Properties.<var>.start` | Character offset of the first occurrence (used only for sort order). |
| `Properties.<var>.auto_size` | Inferred font size from surrounding text, or `null`. |
| `Properties.<var>.font_name` | Nearest run's font name, or `null`. |

**Ordering:** the order of list entries is preserved and shown in the dropdown. Drag & drop in the dropdown re-saves this file.

**Deletion semantics:** when a template is deleted, WordCloner also:
- removes the `.docx` file **unless** another record shares the same `address` (multiple names for one file);
- removes the name from `quick_access` in settings.

---

## 5. 🔬 The Placeholder Pipeline

This is the heart of the project. Both GUI and CLI use it.

### 5.1 Extraction (`extract_placeholders`)

Input: path to a `.docx`.

Steps:

1. Load the document with `python-docx`.
2. Collect paragraphs **in document reading order**:
   - Body paragraphs (`w:p` children of `w:body`).
   - Paragraphs inside top-level tables (`w:tbl`).
   - Nested tables are **not** traversed (limitation).
3. Join paragraph texts with `\n`. Normalize curly quotes (`“ ” „ ‟ ″`) to `"`.
4. Normalize the wrapper too, then find all occurrences of the wrapper.
5. If the count is **odd**, raise `WrapperCountError`.
6. For each pair `(open_pos, close_pos)`:
   - Extract `key = text[open_pos+wlen : close_pos].strip()`.
   - Skip if empty or already seen.
   - Find the paragraph index `pi` whose start ≤ `open_pos`.
   - Compute `local_start` and `local_end` within that paragraph.
   - Rebuild the paragraph's character→run map.
   - Call `compute_auto_font(...)` to infer the size.
   - Find the nearest run and grab its `font.name`.
   - Store `{ "start": term_start, "auto_size": ..., "font_name": ... }`.

Result: `dict[str, dict]` in **first-appearance order**.

### 5.2 `compute_auto_font` — how it works

For the placeholder at `[ph_start, ph_end)`:

- Walk **backwards** word by word; for each word take the run of its first character and record `run.font.size.pt`.
- Walk **forwards** the same way.
- Compare `before[d]` and `after[d]` for `d = 0, 1, 2, …`. The first pair with `abs(b - a) < 0.01` wins.
- If no matched pair, return the **most common rounded size** from all collected sizes.
- If nothing was collected, return `None` (means "auto" → apply the reference run's size).

### 5.3 Replacement (`replace_placeholders_in_docx`)

For each paragraph:

1. Build a **character → (run_index, char_index)** map.
2. Reconstruct normalized text.
3. Find non-overlapping `(start, end, key)` ranges **only for keys present in `replacements`**.
4. Compute per-range `auto_sizes` (from override or auto) and `auto_refs` (nearest run).
5. Build a segment list:
   - Segments of **original** text (preserve run formatting).
   - Segments of **replacement** text (apply `font_overrides` and reference font).
6. Remove all original runs.
7. Append new runs from the segment list.

**Preservation details for original segments:**
- `bold`, `italic`, `underline`
- `font.size`, `font.name`
- `font.color.rgb`

**Replacement segments:**
- Copy `w:rFonts` from reference run (via `_copy_run_fonts`).
- Apply `size_pt` if not `None`.
- Apply `bold` and `italic` (always set, even if `False` — this can override existing styling).
- Apply `color_hex` if provided.

### 5.4 Atomic write (`_save_docx_atomically`)

1. Create parent directory if needed.
2. Refuse to overwrite if `overwrite=False` and destination exists.
3. `tempfile.mkstemp` in the destination folder.
4. Write to temp path.
5. `os.replace(temp, dst)` — atomic on Windows NTFS and POSIX.
6. Delete temp file in `finally` if it still exists.

---

## 6. 🖥️ The GUI Layer

### 6.1 Page stack

`MainWindow.stack` is a `QStackedWidget` containing:

| Index | Widget | Purpose |
|-------|--------|---------|
| 0 | `HomePage` | Dropdown + Quick Access + settings/guide buttons. |
| 1 | `AddLetterPage` | Register a new template. |
| 2 | `EditLetterPage` | Fill variables and issue one document. |
| 3 | `BulkIssuePage` | Excel-driven batch issuance. |
| 4 | `SettingsPage` | Language, output dir, wrapper. |

### 6.2 Navigation signals

- `HomePage.settings_requested` → `MainWindow._go_to_settings`
- `HomePage.quick_template_selected` → `_on_template_selected`
- `HomePage.dropdown.template_selected` → `_on_template_selected`
- `HomePage.dropdown.add_requested` → `_go_to_add`
- `AddLetterPage.saved` → `_on_letter_saved` → `reload_templates`
- `EditLetterPage.batch_requested` → `_open_bulk_issue`
- `BulkIssuePage.back_requested` → `_return_to_issue_page`
- `SettingsPage.saved` → `_on_settings_saved` → if language changed, `theme_changed` → **restart window**.

### 6.3 Language change requires restart

When language changes, `MainWindow.theme_changed` is emitted, caught by `restart()` inside `gui_main`. The old window is hidden and deleted, settings reloaded, and a new `MainWindow` is created with the same geometry.

**Why restart?** Qt layouts are built with `setLayoutDirection()` at widget construction. Rebuilding is simpler than retranslating every string.

### 6.4 Frameless window and Aero Snap

- Window uses `Qt.FramelessWindowHint`.
- `nativeEvent` handles `WM_NCCALCSIZE` (returns `True, 0`) and `WM_NCHITTEST` (returns edge codes for resizing).
- `_apply_win32_styles` ORs `WS_THICKFRAME`, `WS_CAPTION`, `WS_SYSMENU`, `WS_MINIMIZEBOX`, `WS_MAXIMIZEBOX` into the window style. **This is what enables Windows Snap.**
- DPI scaling is accounted for by dividing mouse coordinates by `devicePixelRatio()`.

### 6.5 Drag & drop within lists

`DragController` manages:
- `start(item, global_pos)` — show ghost, create indicator.
- `move(global_pos)` — update ghost position, compute target index.
- `end()` — hide ghost, reorder items, emit `on_reorder`.

Both `DropdownPopup` and `QuickAccessWidget` use it.

**Gotcha:** the drag indicator is placed using `ITEM_H` pixels — the item widget must be exactly `ITEM_H` tall (48 for popup, 44 for QA).

### 6.6 Color picker localization

Qt's `QColorDialog` is used but its labels ("OK", "Cancel", "HTML") are replaced at runtime:
- Buttons are found via `findChildren(QPushButton)`.
- The label with text `"html"` is renamed to `"HEX:"`.
- The dialog is forced `LeftToRight` layout so colors picker controls stay logical.

---

## 7. ⌨️ The CLI Layer

### 7.1 Argument reference

| Short | Long | Type | Notes |
|-------|------|------|-------|
| `-i` | `--input` | path | **Required.** `.docx` template. |
| `-o` | `--output` | path | Single mode: `.docx`; Excel mode: directory. |
| `-v` | `--values` | list | Values in first-appearance order. |
| `-w` | `--wrapper` | text | Override wrapper for this run. |
| `-x` | `--excel` | path | Excel workbook. |
| | `--header-row` | flag | Skip row 1. |
| `-c` | `--count` | flag | Print variable count and exit. |
| | `--force` | flag | Overwrite single-mode output. |
| | `--gui` | flag | Force GUI (used by desktop shortcut). |

### 7.2 Mode decision tree

```
If --count:                 → count mode (print count, exit 0)
Elif --excel:               → batch mode
Elif -o and -v:             → single issuance
Else:                       → parser.error(...)
```

**Conflicts checked:**
- `--count` cannot be combined with `-o`, `-v`, `-x`, `--header-row`, `--force`.
- `--header-row` without `--excel` → error.
- `--excel` with `-v` or `--force` → error.
- In Excel mode, `-o` must be a folder, not a `.docx`.

### 7.3 Value validation

- `len(values) > len(variable_names)` → error `cli_too_many_values`.
- `len(values) < len(variable_names)` → error `cli_too_few_values`.
- Any value == `""` → error `cli_empty_value` (must use `" "` for blank).
- Output equals input → error `cli_output_is_input`.
- Output exists and no `--force` → error `cli_output_exists`.

### 7.4 Windows console attach

If `stdout` or `stderr` is `None` (typical for `pythonw.exe` or a windowed frozen exe), `_attach_cli_console_if_needed()` calls `AttachConsole(-1)` (parent) or `AllocConsole()` and reopens `CONIN$` / `CONOUT$`.

**Caveat:** if the parent CMD is not waiting (windowed exe launched from Explorer), outputs will be lost. The recommended `.spec` uses `console=True` + `hide_console='hide-late'` so the console appears, is hidden when the GUI starts, but is inherited by CMD.

### 7.5 Exit codes

| Code | Meaning |
|------|---------|
| 0 | Success. |
| 2 | `argparse` error (invalid flags, missing required). |
| 1 | Unhandled exception (should not normally happen). |

---

## 8. 🌍 Internationalization (i18n)

### 8.1 Structure of `T`

Each field is:

```python
app_title = _entry("سامانه صدور نامه", "Letter Issuance System")
```

Which expands to:

```python
{"persian": "سامانه صدور نامه", "english": "Letter Issuance System"}
```

### 8.2 Translation lookup

```python
def tr(item, **kwargs):
    key = LANG_MAP.get(SETTINGS["language"], "persian")
    value = item.get(key) or item.get("persian", "")
    return value.format(**kwargs) if kwargs else value
```

### 8.3 RTL handling

- `_is_rtl()` returns `True` when `language == "fa"`.
- `_dir()` returns `Qt.RightToLeft` or `Qt.LeftToRight`.
- Applied to: the whole app (`app.setLayoutDirection`), dialogs, popups, combos.

### 8.4 Adding a new language (developer task)

1. Add the code to `LANG_MAP` (`"fr": "french"`).
2. Add a `"french"` key to every `_entry` — or make `_entry` accept a third argument.
3. Add the language to `SettingsPage.LANGUAGES`.
4. Update the `language` validation in `load_settings`.
5. Update `HomePage._open_guide` to select the correct video file.

---

## 9. ⚙️ Settings System

- **Source of truth:** `SETTINGS` module-level dict.
- **Persistence:** `save_settings()` writes UTF-8 JSON with `indent=4`.
- **Load:** `load_settings()` resets to defaults, then updates from disk.
- **Validation:** on load, `quick_access` must be a list; `language` must be `"fa"` or `"en"`.

**Auto-save triggers:**
- Settings page save button.
- Quick Access add/remove/reorder.
- Template deletion (to clean `quick_access`).
- Window resize/move (debounced 350 ms via `_window_state_timer`).
- Window maximize/restore.
- On close (`closeEvent`).

**Do not write settings from a non-main thread.** `save_settings` uses `open()` without locking.

---

## 10. 🎨 Icon & Resource System

### 10.1 Where icons come from

`wordcloner_icons.ICON_DIR` resolves to:

1. `Path(sys.executable).parent / "resources" / "icons"` if frozen.
2. `Path(__file__).parent / "resources" / "icons"` if running from source.
3. Fallback: `importlib.resources.files("wordcloner_assets").joinpath("icons")` — used when installed as a wheel with a private asset package.

### 10.2 Rendering

`_render_svg_icon(name, size, color)`:

1. Read `name.svg` as text.
2. Replace `currentColor` with the desired hex color.
3. Build `QSvgRenderer` from a `QByteArray`.
4. Render into a transparent `QPixmap` of the requested size.
5. Wrap in `QIcon`.

Cached by `(name, size, color)` with `lru_cache(maxsize=256)`.

### 10.3 Adding an icon

1. Drop `new-icon.svg` in `resources/icons/`.
2. Add it to `WordCloner.spec` under `datas`.
3. Add a public function in `wordcloner_icons.py`.

---

## 11. 🛣️ Paths & Legacy Migration

### 11.1 Current storage layout (frozen .exe)

```
<install-dir>/                     ← APP_ROOT
├── wordcloner.exe
├── app.ico
├── settings.json                  ← created on first save
├── templates.json                 ← created on first template save
├── Words/                         ← copied template .docx files
├── resources/icons/               ← SVG assets
├── Guide/                         ← MP4 tutorials
└── _internal/                     ← PyInstaller runtime
```

### 11.2 Override via environment variable

```
set WORDCLONER_HOME=C:\WordClonerData
wordcloner.exe
```

All data (`settings.json`, `templates.json`, `Words/`) will be created under that folder.

### 11.3 Migration sources

On every startup, `migrate_legacy_data()` runs and:

1. Copies missing files from `APP_ROOT / "WordCloner"` (legacy project-local).
2. Copies missing files from `~/Documents/WordCloner`.
3. Merges `templates.json` records (existing root-level records win).
4. Rewrites old absolute `address` values to point under the new root.

**Originals are never deleted.** Users must remove them manually.

---

## 12. 📊 Batch Issuance from Excel

### 12.1 Expected sheet layout

| Column 1 | Column 2 | Column 3 | … |
|----------|----------|----------|---|
| output filename (no extension needed) | value for variable #1 | value for variable #2 | … |

- Columns beyond `1 + len(variable_names)` that contain data → error.
- Row 1 is optional header (`--header-row` / checkbox).
- Blank cells → blank replacement.

### 12.2 Cell normalization

- `datetime` → `"YYYY-MM-DD"` or `"YYYY-MM-DD HH:MM:SS"`.
- `date` → `"YYYY-MM-DD"`.
- `float` integers → `"42"` not `"42.0"`.
- Everything else → `str(value)`.
- `None` → `""`.

### 12.3 Filename safety

`_safe_batch_filename`:

- Replaces `<>:"/\|?*` and control chars with `_`.
- Strips trailing dots/spaces.
- Appends `.docx` if missing.
- Prefixes reserved Windows names (`CON`, `PRN`, `AUX`, `NUL`, `COM1-9`, `LPT1-9`) with `_`.
- Deduplicates: `name.docx`, `name_2.docx`, `name_3.docx`, …

### 12.4 Output folder

Default: `<output_dir>/<"صدور جمعی" or "Bulk Issuance">/`.

The folder name is **translated**, meaning a user who switches languages mid-project may end up with two folders.

---

## 13. 🔔 System Tray Integration

- `QSystemTrayIcon` with `app.ico` (or a fallback icon).
- Context menu (rebuilt on any Quick Access change):
  - **Open App**
  - **Quick Access** header + one action per valid template name.
  - **Exit**
- Left click / double click toggles visibility.

**Close behavior:** closing the window hides it (does not quit). Only **Exit** in the tray menu (or a `QApplication.quit()` call) sets `_really_quit = True` and closes for good.

---

## 14. 📦 Build & Packaging

### 14.1 PyInstaller command (from `src/`)

```bat
pyinstaller --noconfirm --clean WordCloner.spec
```

### 14.2 Recommended `.spec` settings

- `console=True` with `hide_console='hide-late'` — allows CLI to work while hiding the console during GUI.
- `datas`:
  - `("resources/icons", "resources/icons")`
  - `("../Practice", "Practice")` — optional.
  - `("Guide", "Guide")` — video files.
  - `("app.ico", ".")`
- `hiddenimports`: `docx`, `openpyxl`, `PySide6.QtSvg`, `win32gui`, `win32con`.
- `excludes`: `tkinter`, `matplotlib`, `scipy` (unless needed).

### 14.3 Distribution folder

PyInstaller's `--onedir` (default) produces:

```
dist/WordCloner/
├── wordcloner.exe
├── app.ico
├── python3.dll, python314.dll, VCRUNTIME140*.dll
├── libcrypto-3.dll, libffi-8.dll, libssl-3.dll
├── _asyncio.pyd, ... (CPython stdlib DLLs)
├── PySide6/      (Qt DLLs, plugins, translations)
├── docx/         (python-docx templates)
├── lxml/         (XML parser)
├── numpy/, PIL/, psutil/, shiboken6/, win32/
├── resources/icons/
└── _internal/
```

**This entire folder must be shipped to the installer.** The Inno Setup script copies `dist\WordCloner\*` recursively.

---

## 15. 📦 Installer (Inno Setup)

### 15.1 Key sections

```iss
[Setup]
ChangesEnvironment=yes
PrivilegesRequired=admin

[Tasks]
Name: "desktopicon"; Description: "..."; Flags: checkedonce

[Files]
Source: "dist\WordCloner\*"; DestDir: "{app}"; Flags: recursesubdirs

[Icons]
Name: "{group}\WordCloner";       Filename: "{app}\wordcloner.exe"
Name: "{autodesktop}\WordCloner"; Filename: "{app}\wordcloner.exe"; Tasks: desktopicon

[Registry]
Root: HKLM; Subkey: "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"; \
    ValueType: expandsz; ValueName: "Path"; \
    ValueData: "{olddata};{app}"; Check: NeedsAddPath('{app}')
```

### 15.2 The `NeedsAddPath` helper

Prevents duplicate PATH entries:

```pascal
function NeedsAddPath(Param: string): boolean;
var OrigPath: string;
begin
  if not RegQueryStringValue(HKLM, 'SYSTEM\...\Environment', 'Path', OrigPath) then
  begin
    Result := True;
    exit;
  end;
  Result := Pos(';' + Param + ';', ';' + OrigPath + ';') = 0;
end;
```

### 15.3 Uninstall cleanliness

`RemoveFromPath` in `CurUninstallStepChanged(usPostUninstall)` strips the install directory from `PATH` so it doesn't leak.

### 15.4 Post-install

The user must **open a new Command Prompt** for the PATH change to take effect. Old CMD windows keep the inherited environment.

---

## 16. ❓ Common User Questions & Answers

This section is **specifically designed to help an AI answer end-user questions** correctly.

### Q1: "How do I create a template?"

Open Word, write the letter, and wrap each variable with `##`:

```
Dear ##customer_name##, your invoice ##invoice_no## is due on ##due_date##.
```

Save as `.docx`. Then in WordCloner: **Add Letter → fill name + drop the file → Confirm**.

### Q2: "Why does WordCloner say 'no variables found'?"

The wrapper is not present, or it's mismatched. Check:
- Are the markers exactly `##`? (default)
- Did you use `##` at **both** sides? Opening and closing must match.
- Did the wrapper change in Settings?

### Q3: "Why 'unmatched placeholder wrappers'?"

The number of `##` tokens is odd. Count them: every opening marker needs a matching closing marker. Example problems:
- `Dear ##name,` — missing closing `##`.
- Two variables but one closing marker missing.

### Q4: "How do I make a variable blank in the CLI?"

Pass a single space inside quotes:

```bat
wordcloner -i template.docx -o out.docx -v "Ali" " " "Operations"
```

Do **not** pass `""` — that triggers the `cli_empty_value` error.

### Q5: "Where are my documents saved?"

Check **Settings → Output Folder**. Default is the user's Desktop. Batch outputs go into `<output_dir>/Bulk Issuance/` (or `صدور جمعی` in Persian).

### Q6: "How do I make the CLI work?"

The Inno Setup installer adds the folder to `PATH`. After installation, **open a new CMD window** and run:

```bat
wordcloner --help
```

### Q7: "Where are templates and settings stored?"

In the same folder as `wordcloner.exe` (or, if `WORDCLONER_HOME` is set, in that folder).

### Q8: "The video guide won't open."

Check that `Guide\Persian_guide.mp4` or `Guide\English_guide.mp4` exists next to `wordcloner.exe`. The correct file is chosen based on the UI language.

### Q9: "How do I add a template to Quick Access?"

Click **⊕** next to "Quick Access" and tick the templates you want. Reorder by dragging the grip icon.

### Q10: "How do I make the output filename unique?"

Batch mode does it automatically: `name.docx`, `name_2.docx`, etc.

---

## 17. 🔌 Common Developer Extension Points

### 17.1 Adding a new setting

1. Add a key to `DEFAULT_SETTINGS` in `WordCloner.py`.
2. Add UI in `SettingsPage` (read the current value in `__init__`).
3. Save it in `SettingsPage._on_save`.
4. Use it wherever needed.

### 17.2 Adding a new placeholder wrapper syntax

The wrapper is just a string. Users can set it in Settings. If you want **multiple** simultaneous wrappers, you'd need to extend `extract_placeholders` and `replace_placeholders_in_docx` — currently only one is supported per run.

### 17.3 Adding a new output format (PDF, ODT)

Add a converter step after `_save_docx_atomically`:

```python
# after producing .docx
subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", str(docx_path)])
```

Best placed inside a new helper `_save_document(output_format, ...)`.

### 17.4 Adding a new page to the GUI

1. Create a `QWidget` subclass.
2. Add it to `MainWindow.stack` in order.
3. Add navigation signals.
4. Update `_go_to_*` methods.

### 17.5 Adding a new icon

See section 10.3.

### 17.6 Supporting nested tables in `extract_placeholders`

Current code only iterates `table.rows[i].cells[j].paragraphs`. For nested tables you'd need a recursive walk over `cell.tables`.

### 17.7 Reading variables from Word headers/footers

Not currently supported. To add it, iterate `doc.sections[*].header.paragraphs` and `.footer.paragraphs` and merge into `_template_paragraphs_in_document_order`.

### 17.8 Adding conditional logic to templates

Not supported. Would require a mini-language parser (e.g. `{{#if condition}}...{{/if}}`). Out of scope for the current engine.

---

## 18. ⚠️ Known Limitations & Gotchas

1. **Nested tables are ignored.** Only top-level table paragraphs are scanned.
2. **Headers and footers are ignored.** Placeholders in headers won't be replaced.
3. **Text boxes and shapes are ignored.** `python-docx` doesn't traverse them.
4. **Placeholders split across runs.** Handled — the code joins runs into a single string before searching.
5. **Placeholders spanning paragraphs.** Not supported. Every placeholder must be inside one paragraph.
6. **Consecutive placeholders with no separator** (`##a####b##`) → the wrapper `##` will be consumed ambiguously. Use a separator.
7. **Bold/italic overrides always set.** If a variable's `bold` checkbox is unchecked, the output run's bold is set to `False` even if surrounding text was bold. This is a design decision.
8. **Automatic font sizing only considers body text.** Table-based templates may see inconsistent results.
9. **Empty `##` pairs** (`####`) are skipped during extraction.
10. **Duplicate variable names** — only the first occurrence's metadata is stored.
11. **Case-sensitive names.** `##name##` ≠ `##Name##`.
12. **Filename collisions in batch mode are auto-resolved** with `_2`, `_3` suffixes — not overwritten.
13. **Output folder name is translated.** Switching UI language mid-project creates a second batch folder.
14. **`WORDCLONER_HOME` changes data location.** Existing data in the install folder is **not** migrated automatically.
15. **`migrate_legacy_data` copies, never moves.** Old folders remain.
16. **Settings save is not atomic.** A crash during `json.dump` could truncate `settings.json`. Recovery: delete the file to reset to defaults.
17. **Window geometry save is debounced 350 ms.** A crash within 350 ms of a resize loses that resize.
18. **Console attach on Windows can only work for the parent process.** If launched from Explorer, CLI output is not visible.
19. **`python-docx` doesn't preserve `sectPr` modifications** made outside the standard API. Complex templates with unusual section properties may lose them.
20. **Right-to-left rendering in `QComboBox` requires the layout direction to be set before contents.** Done in `ComboWithArrow.__init__`.

---

## 19. 🐛 Debugging & Diagnostics

### 19.1 Enable verbose logging

The app prints errors to `sys.stderr` with `[error]`, `[warning]`, `[migration]`, `[icons]` prefixes. Run from a console to see them.

### 19.2 Reproduce a bug in isolation

Call the pipeline directly in Python:

```python
from wordcloner_paths import migrate_legacy_data
from WordCloner import (
    load_settings, extract_placeholders, replace_placeholders_in_docx,
)

migrate_legacy_data()
load_settings()

props = extract_placeholders("sample.docx")
print(props)

replace_placeholders_in_docx(
    "sample.docx", "out.docx",
    {"customer_name": "Ali", "invoice_no": "INV-001"},
    {"customer_name": (14.0, True, False, "#000000")},
)
```

### 19.3 Inspect `templates.json`

Any time a template "doesn't show up", open `templates.json` and check:
- Is the record's `name` non-empty?
- Does `address` point to a `.docx` that exists?
- Does `Properties` contain at least one key?

### 19.4 Common Windows-only bugs

- PATH not updated → check `HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\Path`.
- Tray icon missing → check Windows "Show hidden icons" or the notification area settings.
- Frozen app can't find icons → verify `resources/icons/*.svg` is present next to the `.exe`.

---

## 20. 📚 Glossary

| Term | Meaning |
|------|---------|
| **Template** | A `.docx` file with placeholder variables. |
| **Variable** | Text between two wrapper markers, e.g. `##name##` → key `name`. |
| **Wrapper** | The marker used on both sides of a variable. Default `##`. |
| **Quick Access** | Pinned shortlist of templates on the home page. |
| **Batch issuance** | Generating many documents from an Excel workbook. |
| **`auto_size`** | Inferred font size for a variable, from surrounding text. |
| **Frozen** | A PyInstaller-built `.exe` — `sys.frozen` is `True`. |
| **`APP_ROOT`** | The directory WordCloner treats as its data root. |
| **`WORDCLONER_HOME`** | Environment variable to override `APP_ROOT`. |
| **`WrapperCountError`** | Raised when the number of markers is odd. |
| **`ExcelImportError`** | Raised when a workbook can't be mapped to the template. |

---

## 21. 🎓 How an AI Should Answer Questions

When you receive a question about WordCloner, follow this decision tree:

### 21.1 Classify the question

| Type | Signal words | Where to look |
|------|--------------|---------------|
| **User – how-to** | "how do I", "where is", "why doesn't" | Sections 5, 12, 16 |
| **User – error** | an error message text | Section 16, 18, 19 |
| **Developer – code** | function names, "how to add" | Sections 3, 17 |
| **Developer – build** | "PyInstaller", "installer", "spec" | Sections 14, 15 |
| **Developer – architecture** | "where is X handled" | Sections 2, 3 |

### 21.2 Answer format

Prefer this structure:

1. **Direct answer** — one sentence.
2. **Explanation** — 2–4 sentences referencing the module/function.
3. **Concrete example** — CLI command, code snippet, or step-by-step.
4. **Cross-references** — link to a section in this file.

### 21.3 Things you must NOT do

- ❌ Invent CLI flags that don't exist.
- ❌ Claim headers/footers are processed — they are not.
- ❌ Claim nested tables are traversed — they are not.
- ❌ Suggest editing `.exe` internals — always rebuild from source.
- ❌ Recommend third-party tools unless the user's question explicitly calls for alternatives.
- ❌ Assume the user's locale. Ask whether they want a Persian or English answer if unclear.

### 21.4 Sample prompts the user may send

- "با WordCloner چطور میتوانم ۱۰۰ نامه بسازم؟" → explain Excel batch mode (section 12).
- "چرا متغیرها تشخیص داده نمیشوند؟" → check wrapper, even count (sections 5, 16 Q2–Q3).
- "How do I add a new language?" → section 8.4.
- "How does auto font sizing work?" → section 5.2.
- "Where is `WORDCLONER_HOME` read?" → `wordcloner_paths._application_root`.
- "Can I use `{{ }}` as wrapper?" → yes, via Settings (section 9).
- "Why does bold unchecking remove bold in output?" → design decision, section 18 #7.
- "How do I build the installer?" → sections 14, 15.

### 21.5 If the answer is unknown

Say so explicitly. Suggest the user:
1. Run the app from a console to see `[error]` logs.
2. Open `templates.json` / `settings.json` to inspect state.
3. Post a minimal `.docx` template that reproduces the issue.

Do not speculate about undocumented behavior.

---

## 📎 Appendix A — Quick Reference Card

**Run GUI:**
```bat
wordcloner
wordcloner --gui
```

**Count variables:**
```bat
wordcloner -i template.docx --count
```

**Single issuance:**
```bat
wordcloner -i template.docx -o out.docx -v "Ali" " " "Ops" --force
```

**Batch issuance:**
```bat
wordcloner -i template.docx --excel data.xlsx --header-row
```

**Change wrapper on the fly:**
```bat
wordcloner -i template.docx --count -w "{{ }}"
```

**Override data root:**
```bat
set WORDCLONER_HOME=D:\WordClonerData
wordcloner
```

**Files you should rarely need to edit:**
- `wordcloner_translations.py` — only when adding strings.
- `wordcloner_icons.py` — only when adding icons.
- `wordcloner_paths.py` — only when changing storage layout.

**Files you will edit often:**
- `WordCloner.py` — almost everything.
- `WordCloner.spec` — when adding/removing bundled data.

---

## 📎 Appendix B — Signal Flow Diagram (GUI)

```
HomePage
├─ gear_btn.clicked ───► settings_requested ───► MainWindow._go_to_settings
├─ guide_btn.clicked ──► _open_guide (internal)
├─ dropdown.add_requested ──► MainWindow._go_to_add
├─ dropdown.template_selected ─► _on_template_selected
├─ dropdown.template_deleted ──► _on_template_deleted
├─ quick_access.template_selected ─► quick_template_selected ─► _on_template_selected
└─ quick_access.quick_access_changed ─► MainWindow._refresh_tray_menu

AddLetterPage
├─ back_requested ──► MainWindow._go_to_home
└─ saved          ──► MainWindow._on_letter_saved ─► reload_templates

EditLetterPage
├─ back_requested  ──► MainWindow._go_to_home
└─ batch_requested ──► MainWindow._open_bulk_issue ─► BulkIssuePage.load_template

BulkIssuePage
└─ back_requested ──► MainWindow._return_to_issue_page

SettingsPage
├─ back_requested ──► MainWindow._go_to_home
└─ saved          ──► MainWindow._on_settings_saved
                       ├─ language unchanged → _go_to_home
                       └─ language changed   → theme_changed → restart()
```

---

## 📎 Appendix C — Example Error Messages (verbatim)

| Key | Persian | English |
|-----|---------|---------|
| `err_no_vars` | در این قالب هیچ متغیری پیدا نشد. | No variables were found in this template. |
| `err_wrapper_count` | تعداد نشانگرهای متغیر با هم مطابقت ندارد | Unmatched Placeholder Wrappers |
| `err_name_exists` | قالبی با نام «{name}» از پیش ثبت شده است. | A template named "{name}" is already registered. |
| `cli_empty_value` | برای مقدار خالی از یک فاصلهٔ داخل نقلقول استفاده کنید (" ")، نه رشتهٔ خالی (""). | Use a quoted single space (" "), not an empty string (""), for a blank value. |
| `cli_too_few_values` | تعداد مقدارهای واردشده ({actual}) از تعداد متغیرهای قالب ({expected}) کمتر است. | Too few values: supplied {actual}, but the template has {expected} unique variable(s). |
| `batch_extra_columns` | سطر {row} در ستون {column} دادهٔ اضافه دارد. | Row {row} has extra data in column {column}. |

---

## 📎 Appendix D — Versioning Notes

- **AppUserModelID:** `com.wordcloner.letterissuance.1.0` — set via `SetCurrentProcessExplicitAppUserModelID`. Changing this breaks taskbar pinning.
- **Installer name:** `WordCloner-setup-63bit.exe`.
- **GUID in Inno Setup:** generate a fresh one for each release branch; keeping it constant across versions enables "upgrade" installs.
- **Python version compatibility:** tested with **3.10+**. Uses `str.removeprefix`-style APIs only where present; if downgrading, check `Path.unlink(missing_ok=True)` (3.8+).

---

## ✅ Summary Checklist for an AI Assistant

Before answering, verify:

- [ ] I read this entire file.
- [ ] I identified whether the user is a **user** or **developer**.
- [ ] I identified GUI vs CLI context.
- [ ] I checked section 18 for the relevant limitation.
- [ ] My answer cites a **module** and **function**.
- [ ] I provided a runnable example if applicable.
- [ ] I did not invent features.
- [ ] I offered to clarify ambiguity.

**End of AI-Readme.md**