"""
================================================================================
 WordCloner — Letter Issuance System
================================================================================
 A PySide6 desktop application for Windows that lets users register Word
 document templates containing placeholders, fill in values, and export
 finished letters with the placeholders replaced.

 Author's intent (for future maintainers / AI agents):
   - This project is a single-file application. All UI, business logic,
     persistence, and file operations live in this file.
   - The code is intentionally structured as small, single-responsibility
     classes so future changes remain local.
   - The theme system is data-driven (see THEMES) and hot-swappable at
     runtime by rebuilding the main window.

================================================================================
 REQUIREMENTS
================================================================================
   Python 3.9+
   pip install PySide6 python-docx

================================================================================
 FILE LAYOUT / DIRECTORY STRUCTURE
================================================================================
   %USERPROFILE%/Documents/WordCloner/
       ├── Words/                     # copied .docx template files
       ├── templates.json             # list of template metadata
       └── settings.json              # user settings (theme, output dir, ...)

   Paths are defined near the top of the file as module-level constants:
       DOCUMENTS_DIR   = Path.home() / "Documents"
       WORDCLONER_DIR  = DOCUMENTS_DIR / "WordCloner"
       WORDS_DIR       = WORDCLONER_DIR / "Words"
       TEMPLATES_JSON  = WORDCLONER_DIR / "templates.json"
       SETTINGS_JSON   = WORDCLONER_DIR / "settings.json"

================================================================================
 DATA SCHEMAS
================================================================================

   --- templates.json ---
   A JSON array of objects. Each object represents one template:
       [
         {
           "name":  "Letter of Article Acceptance",   # user-chosen name
           "address": "C:\\...\\Words\\Accepted.docx",  # absolute path to .docx
           "Properties": {
               "Author Name": 145,       # key = placeholder text
               "Article Title": 210      # value = char index of placeholder start
           }
         },
         ...
       ]

   --- settings.json ---
   A single JSON object:
       {
         "theme": "green",                              # key in THEMES dict
         "output_dir": "C:\\Users\\...\\Desktop",       # where finished .docx go
         "placeholder_wrapper": "\"\""                  # wrapper around variables
       }

================================================================================
 GLOBAL STATE
================================================================================
   SETTINGS           : dict   - Runtime copy of settings.json content.
                                 Always access via SETTINGS["key"], never by
                                 re-reading the file (except on load/restart).
                                 Written only by save_settings().

   THEMES             : dict   - Static palette definitions. Keys are theme
                                 ids ("green","blue","purple"). Values are
                                 dicts with fields: label, dark, primary,
                                 light, hover, bg, border, muted, accent1,
                                 accent2.

   C_*                : module-level color strings that mirror the *current*
                                 theme. Refreshed by reload_theme_colors().
                                 These are read directly by icons, stylesheets,
                                 and paintEvent methods. Because they are read
                                 at widget construction time, changing the theme
                                 requires rebuilding the whole UI (see main()).

   HANDLING THEME CHANGES:
     - SettingsPage writes new theme to SETTINGS and emits saved.
     - MainWindow._on_settings_saved() compares dark color of old/new theme.
     - If changed, MainWindow.theme_changed is emitted.
     - main()'s restart() slot destroys the old MainWindow, reloads colors,
       and creates a fresh MainWindow. This is the ONLY correct way to
       re-theme, since stylesheets are baked into widgets.

================================================================================
 PLACEHOLDER EXTRACTION & REPLACEMENT
================================================================================
   The wrapper string (SETTINGS["placeholder_wrapper"], default '""') is the
   delimiter around variable names in the Word document.

   Example with default wrapper:
       Hello ""Author Name"", welcome.
       → extracts key "Author Name" at index 7 (after the wrapper).

   IMPORTANT: Word often auto-corrects straight quotes (") to curly quotes
   (“ ”). normalize_quotes() converts them to straight quotes so extraction
   is uniform. Length is preserved, so character indices remain valid.

   extract_placeholders():
       - Opens .docx via python-docx
       - Concatenates all paragraph texts with "\n"
       - Finds ALL occurrences of the wrapper via str.find() in a loop
       - Pairs them positionally: (positions[0], positions[1]), (positions[2],
         positions[3]), ... — this correctly handles the case where two
         placeholders are adjacent, unlike a naive regex.
       - Stores {key: char_index_of_term_start}. If a key repeats, the FIRST
         occurrence wins.

   replace_placeholders_in_docx():
       - For each paragraph (top-level + inside tables):
           * Flatten runs into a single string + a char→(run_idx, offset) map
           * Find all wrapper pairs; skip ones whose key isn't in replacements
           * Build a list of "segments":
               - (text, original_run, is_replacement=False)  → copy formatting
               - (value, (size_pt, bold, italic), True)      → apply override
           * Delete original runs; append new runs from segments
       - Auto font size (when user chose "auto"):
           * Walk outward from placeholder, comparing word-by-word:
               depth 0: (word_before_0, word_after_0)
               depth 1: (word_before_1, word_after_1) ...
           * If at any depth both words have the SAME size → use it.
           * If one side runs out → fall back to the most common size among
             all pairs checked so far.
           * If no words exist → None (inherit paragraph style).
           * See compute_auto_font() for exact implementation.

================================================================================
 CLASS HIERARCHY & RESPONSIBILITIES
================================================================================

   ThemedDialog (QDialog)
     A frameless, modal dialog styled to match the app. Used EVERYWHERE instead
     of QMessageBox.
     Static helpers:
       .info(parent, title, msg)           → OK only
       .warning(parent, title, msg)        → OK only
       .error(parent, title, msg)          → OK only
       .confirm(parent, title, msg)        → returns bool (OK/Cancel)
       .confirm_delete(parent, title, msg) → returns bool, red trash icon
     The dialog auto-centers over its parent in showEvent().
     Kinds ("info","warning","error","delete","question") select the icon and
     accent color of icon_status().

   TitleBar (QFrame)
     Custom title bar replacing the OS chrome (window is Frameless).
     - Holds app icon + title + min/max/close buttons.
     - Implements drag-to-move and double-click-to-maximize by manipulating
       the parent QMainWindow (self._win).
     - Has top-left and top-right CSS border-radius (14px) matching Container.

   Container (QWidget)
     Root visual surface of MainWindow's content, placed inside a 16-px margin
     that exposes the drop shadow applied in MainWindow.
     paintEvent():
       - Fills a rounded rect with C_BG
       - Clips to that shape
       - Draws two decorative green curves at the bottom-left and one faint
         arc at the top-right (colors from current theme accents)
       - Strokes the border in C_BORDER
     Note: The Container's paintEvent is *behind* its children, so widgets
     placed in its layout render on top of the decorations.

   DropdownButton (QFrame)
     The white rounded box labeled "Select template". Emits clicked on press.
     Owns its own stylesheet (hover changes border to C_GREEN).

   PopupItem (QFrame)
     A single row inside DropdownPopup. Two flavors:
       - is_add=True       → green circle "+" icon, bold text, light-green bg
       - show_delete=True  → reveals a small trash QPushButton on hover
     Hover show/hide of the delete button is done via enterEvent/leaveEvent
     on the item; the actual click on the delete button is filtered inside
     mousePressEvent so the row's clicked signal is NOT emitted when the user
     clicks the trash icon.

   DropdownPopup (QWidget, Qt.Popup)
     The floating list of templates + "Add new letter" entry.
     - Qt.Popup flag makes it auto-close on outside click.
     - Emits:
         add_clicked              → user clicked "Add new"
         template_selected(dict)  → user clicked a template row
         template_deleted(dict)   → user clicked the trash on a row

   TemplateDropdown (QWidget)
     Composite: contains a DropdownButton and manages the popup lifecycle.
     - set_templates(list) stores the current list.
     - _toggle() opens/closes the popup, position it under the button.
     - Relays popup signals outward as:
         add_requested, template_selected(dict), template_deleted(dict)
     Re-positions popup at button.mapToGlobal(0, button.height()).

   DropZone (QFrame)
     Drag-and-drop target for a single .docx file.
     - Accepts dragEnter only if exactly one local file ending in ".docx".
     - Applies a "hover" stylesheet while dragging over.
     - Clicking anywhere triggers QFileDialog.getOpenFileName.
     - On success calls _set_file() which updates the label and emits
       file_selected(path).
     - reset() clears the file and label. Called before entering AddLetterPage.

   RoundToggle (QPushButton, checkable)
     Small circular toggle used for Bold ("B") and Italic ("I") per variable.
     Visual state driven entirely by :checked pseudo-class in stylesheet.

   FontSizeCombo (QComboBox, editable)
     Pre-populated with FONT_SIZES = ["auto","8","9",...,"72"]. The user can
     also type a custom size. Caller must parse currentText() manually:
       "" or "auto" → None (use auto algorithm)
       float-parseable → numeric pt size

   HomePage (QWidget)
     Landing page. Layout (RTL): gear button on the right, then a centered
     card with app icon, big title "Letter Issuance", and TemplateDropdown.
     Signals:
       settings_requested → gear click

   AddLetterPage (QWidget)
     Register a new template.
     - name_input (QLineEdit)  — user-facing template name
     - drop_zone (DropZone)    — source .docx
     - confirm_btn             — validates + copies + extracts + persists
     Signals:
       back_requested → top-left back button
       saved          → after successful save
     Validation rules in _on_confirm:
       * name must be non-empty, must not collide with an existing template
       * a file must be selected
       * python-docx must be importable
     Storage flow:
       1) Copy source .docx into WORDS_DIR (auto-suffix _1, _2 on collision)
       2) extract_placeholders() on the COPY
       3) Append {"name", "address", "Properties"} to templates.json
       4) Show success ThemedDialog, reset form, emit saved

   EditLetterPage (QWidget)
     Fill in variable values and export.
     - filename_input  — target filename (sanitized; .docx appended if missing)
     - scroll          — QScrollArea containing per-variable field rows
     - confirm_btn     — bottom-RIGHT (in RTL this means the last item after
                         an addStretch())
     Signals:
       back_requested → top-left back button; also emitted after successful
                        export to return to HomePage.

     Per-variable row (built by _build_field):
       Row 1 (horizontal):
         [label key ........] [FontSizeCombo] [B RoundToggle] [I RoundToggle]
       Row 2:
         [QLineEdit for value]

     self._fields : {key -> {"input", "size", "bold", "italic"}}
     load_template(tpl):
       - Sets title and default filename
       - Clears previous rows (except the trailing stretch)
       - Sorts Properties by start index so the UI order matches document order
       - Creates one row per key
     _on_confirm():
       - Sanitizes filename, resolves output_dir from SETTINGS
       - Confirms overwrite via ThemedDialog.confirm
       - Builds replacements dict and font_overrides dict
       - Calls replace_placeholders_in_docx(src, out, repl, overrides)
       - On success: info dialog + back_requested

   SettingsPage (QWidget)
     Edits SETTINGS and writes settings.json.
     Fields:
       theme_combo    — maps theme key via userData
       output_input   — free text + "Browse…" button
       wrapper_input  — text (default '""')
     Signals:
       back_requested, saved
     Note: MainWindow._go_to_settings() destroys and recreates this page each
     time to guarantee fresh values from SETTINGS.

   MainWindow (QMainWindow)
     The single top-level window. Owns:
       - Container (central widget) with custom shadow
       - TitleBar
       - QStackedWidget with 4 pages: HomePage, AddLetterPage, EditLetterPage,
         SettingsPage
     Wiring (in __init__):
       home.dropdown.add_requested       → _go_to_add
       home.dropdown.template_selected   → _on_template_selected
       home.dropdown.template_deleted    → _on_template_deleted
       home.settings_requested           → _go_to_settings
       add.back_requested                → _go_to_home
       add.saved                         → _on_letter_saved (reload + home)
       edit.back_requested               → _go_to_home
       settings.back_requested           → _go_to_home
       settings.saved                    → _on_settings_saved (theme diff check)
     Also overrides changeEvent to drop the 16-px shadow margin when maximized.

   main()
     - load_settings() + reload_theme_colors()
     - Creates QApplication, sets RTL direction, picks a reasonable font
     - Defines restart() closure: destroys current window, reloads settings
       and colors, creates a fresh MainWindow. Used to apply theme changes.
     - create_window() builds and shows the first MainWindow, wiring
       theme_changed → restart.

================================================================================
 KEY FUNCTION REFERENCE
================================================================================

   load_settings() / save_settings()
     Read/write settings.json into/from the SETTINGS global.

   reload_theme_colors()
     Copies the active theme's colors into the module-level C_* globals.
     MUST be called after load_settings() and before constructing any widget.

   get_theme() → dict
     Returns the full active theme dict (used by icon_status).

   extract_placeholders(docx_path) → dict[str, int]
     See section PLACEHOLDER EXTRACTION.

   compute_auto_font(runs, char_map, text, ph_start, ph_end) → float | None
     Implements the "auto" font-size algorithm described above. Returns
     None if no neighbors had explicit sizes.

   replace_placeholders_in_docx(src, dst, replacements, font_overrides)
     replacements   : {key: value_str}
     font_overrides : {key: (size_pt_or_None, bold_bool, italic_bool)}
     See PLACEHOLDER EXTRACTION for algorithm details.

   normalize_quotes(text) → str
     Straightens Word's curly quotes. Length-preserving.

   All icon_*() functions
     Pure functions returning QIcon drawn via QPainter. They accept size and
     optional color; colors default to the current theme's C_* globals.
     Because they read C_* at call time, they must be re-invoked after a
     theme change (which the full window rebuild accomplishes).

================================================================================
 EXTENSION GUIDELINES (for future AI contributors)
================================================================================
   1) Adding a new page:
      - Subclass QWidget with a "transparent" stylesheet background.
      - Instantiate it in MainWindow.__init__, add to self.stack.
      - Wire navigation via signals (never call stack.setCurrentWidget across
        pages directly; always go through a MainWindow slot).

   2) Adding a new setting:
      - Add the default to DEFAULT_SETTINGS.
      - Add a widget to SettingsPage; read it in _on_save().
      - Access at runtime via SETTINGS["key"].

   3) Adding a new theme:
      - Add a key to THEMES with the same field set.
      - No other changes needed — the combo populates from THEMES.

   4) Changing placeholder delimiter:
      - Do NOT hardcode '""' anywhere. Always use _get_wrapper().

   5) Changing the auto-font heuristic:
      - Only edit compute_auto_font(). Its contract is:
          inputs  = paragraph runs + char map + normalized text + indices
          output  = float pt size OR None (inherit)
      - The caller assumes the return value is directly usable by Pt().

   6) Changing document traversal:
      - Currently paragraphs and table-cell paragraphs are processed.
      - Headers/footers/sections are NOT processed. If needed, add them in
        replace_placeholders_in_docx by iterating doc.sections and
        doc.sections[i].header/.footer paragraphs.

   7) Changing a widget's colors dynamically:
      - Prefer full-window rebuild via main.restart() as done for themes.
      - Local recoloring is possible but you must also refresh any cached
        QIcon objects (they bake QColor at creation time).

   8) Error handling policy:
      - Never use QMessageBox. Always use ThemedDialog.*
      - Catch broad Exception in user-triggered operations and surface a
        ThemedDialog.error with the exception message.

   9) Persistence rules:
      - templates.json is read/written only through load_templates_raw,
        load_templates, save_templates.
      - settings.json only through load_settings / save_settings.

================================================================================
 KNOWN LIMITATIONS / TODO
================================================================================
   - When a paragraph contains MULTIPLE different runs with distinct
     formatting (e.g., mixed bold/italic mid-paragraph), the replacement
     routine rebuilds runs and copies properties from the original run that
     contained the first character of each output segment. Richly formatted
     runs that were merged during flattening lose sub-run boundaries.
   - Headers/footers and footnotes are not processed.
   - Images inside paragraphs are preserved because we only manipulate text
     runs; but any inline image anchored to a run that gets removed could be
     lost. Test with image-heavy templates.
   - Theme change rebuilds the window; unsaved form state is lost.
   - The auto font size uses only the run's explicit `font.size`. It does not
     fall back to the paragraph style or document defaults, so templates that
     rely on paragraph-level sizes may return None and inherit naturally.
================================================================================
"""

import sys
import json
import shutil
import math
import base64
from copy import deepcopy
from pathlib import Path
from collections import Counter

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFrame, QGraphicsDropShadowEffect,
    QStackedWidget, QLineEdit, QFileDialog, QDialog, QComboBox,
    QScrollArea
)
from PySide6.QtCore import Qt, QPoint, Signal, QSize, QRectF, QEvent
from PySide6.QtGui import (
    QIcon, QPixmap, QPainter, QColor, QFont, QPen, QBrush, QPainterPath,
    QGuiApplication
)

try:
    from docx import Document
    from docx.shared import Pt
    from docx.oxml.ns import qn
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False
    print("[هشدار] python-docx نصب نیست → pip install python-docx")


# ═════════════════════════════════════════════════════════════════
#  مسیرها
# ═════════════════════════════════════════════════════════════════

DOCUMENTS_DIR  = Path.home() / "Documents"
WORDCLONER_DIR = DOCUMENTS_DIR / "WordCloner"
WORDS_DIR      = WORDCLONER_DIR / "Words"
TEMPLATES_JSON = WORDCLONER_DIR / "templates.json"
SETTINGS_JSON  = WORDCLONER_DIR / "settings.json"


# ═════════════════════════════════════════════════════════════════
#  تنظیمات
# ═════════════════════════════════════════════════════════════════

DEFAULT_SETTINGS = {
    "theme": "green",
    "output_dir": str(Path.home() / "Desktop"),
    "placeholder_wrapper": '""',
}
SETTINGS = dict(DEFAULT_SETTINGS)


def load_settings():
    global SETTINGS
    SETTINGS = dict(DEFAULT_SETTINGS)
    if SETTINGS_JSON.exists():
        try:
            with open(SETTINGS_JSON, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                SETTINGS.update(data)
        except Exception as e:
            print(f"[خطا] خواندن settings: {e}")


def save_settings():
    WORDCLONER_DIR.mkdir(parents=True, exist_ok=True)
    try:
        with open(SETTINGS_JSON, "w", encoding="utf-8") as f:
            json.dump(SETTINGS, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"[خطا] ذخیره settings: {e}")


# ═════════════════════════════════════════════════════════════════
#  تم‌ها
# ═════════════════════════════════════════════════════════════════

THEMES = {
    "green": {
        "label": "سبز",
        "dark": "#2E5A4C", "primary": "#3A7A5E",
        "light": "#E8F0EC", "hover": "#F0F5F2",
        "bg": "#F4F7F2", "border": "#D0DDD5", "muted": "#7A8F86",
        "accent1": "#8FB5A2", "accent2": "#E6EEE8",
    },
    "blue": {
        "label": "آبی",
        "dark": "#1E3A5F", "primary": "#2E6FA5",
        "light": "#E5EEF7", "hover": "#EFF5FA",
        "bg": "#F4F7FA", "border": "#CFDBE6", "muted": "#7A8A9E",
        "accent1": "#88A9C9", "accent2": "#E5EBF2",
    },
    "purple": {
        "label": "بنفش",
        "dark": "#4A2E6A", "primary": "#6A46A0",
        "light": "#EDE5F5", "hover": "#F4EFF9",
        "bg": "#F7F4FA", "border": "#D9CFE5", "muted": "#8A7E9E",
        "accent1": "#A890C9", "accent2": "#EEE9F4",
    },
}


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
C_ORANGE      = "#E6A23C"


def reload_theme_colors():
    global C_DARK_GREEN, C_GREEN, C_LIGHT_GREEN, C_HOVER
    global C_BG, C_BORDER, C_MUTED, C_ACCENT1, C_ACCENT2
    t = THEMES.get(SETTINGS.get("theme", "green"), THEMES["green"])
    C_DARK_GREEN  = t["dark"]
    C_GREEN       = t["primary"]
    C_LIGHT_GREEN = t["light"]
    C_HOVER       = t["hover"]
    C_BG          = t["bg"]
    C_BORDER      = t["border"]
    C_MUTED       = t["muted"]
    C_ACCENT1     = t["accent1"]
    C_ACCENT2     = t["accent2"]


def get_theme():
    return THEMES.get(SETTINGS.get("theme", "green"), THEMES["green"])


def _svg_arrow_uri(color_hex):
    """فلش رو به پایین واضح (SVG base64)."""
    svg = (
        "<svg xmlns='http://www.w3.org/2000/svg' width='14' height='14' "
        "viewBox='0 0 14 14'>"
        f"<path d='M3.5 5.2 L7 8.7 L10.5 5.2' stroke='{color_hex}' "
        "stroke-width='2' fill='none' "
        "stroke-linecap='round' stroke-linejoin='round'/>"
        "</svg>"
    )
    b64 = base64.b64encode(svg.encode("utf-8")).decode("ascii")
    return f"url(data:image/svg+xml;base64,{b64})"


# ═════════════════════════════════════════════════════════════════
#  آیکون‌سازها
# ═════════════════════════════════════════════════════════════════

def icon_document(size=22, color=None):
    color = color or C_DARK_GREEN
    pm = QPixmap(size, size); pm.fill(Qt.transparent)
    p = QPainter(pm); p.setRenderHint(QPainter.Antialiasing)
    s = size
    pen = QPen(QColor(color), max(1.4, s * 0.085))
    pen.setJoinStyle(Qt.RoundJoin); pen.setCapStyle(Qt.RoundCap)
    p.setPen(pen); p.setBrush(Qt.NoBrush)
    path = QPainterPath()
    path.moveTo(s * 0.22, s * 0.10); path.lineTo(s * 0.60, s * 0.10)
    path.lineTo(s * 0.80, s * 0.30); path.lineTo(s * 0.80, s * 0.90)
    path.lineTo(s * 0.22, s * 0.90); path.closeSubpath()
    p.drawPath(path)
    p.drawLine(int(s * 0.60), int(s * 0.10), int(s * 0.60), int(s * 0.30))
    p.drawLine(int(s * 0.60), int(s * 0.30), int(s * 0.80), int(s * 0.30))
    p.drawLine(int(s * 0.34), int(s * 0.46), int(s * 0.68), int(s * 0.46))
    p.drawLine(int(s * 0.34), int(s * 0.60), int(s * 0.68), int(s * 0.60))
    p.drawLine(int(s * 0.34), int(s * 0.74), int(s * 0.55), int(s * 0.74))
    p.end(); return QIcon(pm)


def icon_document_big(size=88, color=None):
    color = color or C_DARK_GREEN
    pm = QPixmap(size, size); pm.fill(Qt.transparent)
    p = QPainter(pm); p.setRenderHint(QPainter.Antialiasing)
    s = size
    pen = QPen(QColor(color), s * 0.045)
    pen.setJoinStyle(Qt.RoundJoin); pen.setCapStyle(Qt.RoundCap)
    p.setPen(pen); p.setBrush(Qt.NoBrush)
    path = QPainterPath()
    path.moveTo(s * 0.18, s * 0.06); path.lineTo(s * 0.60, s * 0.06)
    path.lineTo(s * 0.80, s * 0.26); path.lineTo(s * 0.80, s * 0.74)
    path.lineTo(s * 0.18, s * 0.74); path.closeSubpath()
    p.drawPath(path)
    p.drawLine(int(s * 0.60), int(s * 0.06), int(s * 0.60), int(s * 0.26))
    p.drawLine(int(s * 0.60), int(s * 0.26), int(s * 0.80), int(s * 0.26))
    p.drawLine(int(s * 0.30), int(s * 0.38), int(s * 0.68), int(s * 0.38))
    p.drawLine(int(s * 0.30), int(s * 0.52), int(s * 0.68), int(s * 0.52))
    p.drawLine(int(s * 0.30), int(s * 0.64), int(s * 0.52), int(s * 0.64))
    cx, cy, r = s * 0.76, s * 0.78, s * 0.20
    p.setPen(Qt.NoPen); p.setBrush(QBrush(QColor(color)))
    p.drawEllipse(QPoint(int(cx), int(cy)), int(r), int(r))
    pen2 = QPen(QColor("#FFFFFF"), s * 0.045); pen2.setCapStyle(Qt.RoundCap)
    p.setPen(pen2)
    p.drawLine(int(cx - r * 0.5), int(cy), int(cx + r * 0.5), int(cy))
    p.drawLine(int(cx), int(cy - r * 0.5), int(cx), int(cy + r * 0.5))
    p.end(); return QIcon(pm)


def icon_plus_circle(size=28, fg="#FFFFFF", bg=None):
    bg = bg or C_GREEN
    pm = QPixmap(size, size); pm.fill(Qt.transparent)
    p = QPainter(pm); p.setRenderHint(QPainter.Antialiasing)
    s = size
    p.setPen(Qt.NoPen); p.setBrush(QBrush(QColor(bg)))
    p.drawEllipse(0, 0, s, s)
    pen = QPen(QColor(fg), s * 0.10); pen.setCapStyle(Qt.RoundCap)
    p.setPen(pen)
    p.drawLine(int(s * 0.28), int(s * 0.5), int(s * 0.72), int(s * 0.5))
    p.drawLine(int(s * 0.5), int(s * 0.28), int(s * 0.5), int(s * 0.72))
    p.end(); return QIcon(pm)


def icon_trash(size=18, color=None):
    color = color or C_MUTED
    pm = QPixmap(size, size); pm.fill(Qt.transparent)
    p = QPainter(pm); p.setRenderHint(QPainter.Antialiasing)
    s = size
    pen = QPen(QColor(color), s * 0.085)
    pen.setCapStyle(Qt.RoundCap); pen.setJoinStyle(Qt.RoundJoin)
    p.setPen(pen); p.setBrush(Qt.NoBrush)
    p.drawLine(int(s * 0.15), int(s * 0.28), int(s * 0.85), int(s * 0.28))
    p.drawLine(int(s * 0.38), int(s * 0.28), int(s * 0.38), int(s * 0.16))
    p.drawLine(int(s * 0.62), int(s * 0.28), int(s * 0.62), int(s * 0.16))
    p.drawLine(int(s * 0.38), int(s * 0.16), int(s * 0.62), int(s * 0.16))
    path = QPainterPath()
    path.moveTo(s * 0.28, s * 0.28); path.lineTo(s * 0.32, s * 0.86)
    path.lineTo(s * 0.68, s * 0.86); path.lineTo(s * 0.72, s * 0.28)
    p.drawPath(path)
    p.drawLine(int(s * 0.44), int(s * 0.42), int(s * 0.44), int(s * 0.72))
    p.drawLine(int(s * 0.56), int(s * 0.42), int(s * 0.56), int(s * 0.72))
    p.end(); return QIcon(pm)


def icon_chevron_down(size=16, color=None):
    color = color or C_DARK_GREEN
    pm = QPixmap(size, size); pm.fill(Qt.transparent)
    p = QPainter(pm); p.setRenderHint(QPainter.Antialiasing)
    s = size
    pen = QPen(QColor(color), s * 0.14)
    pen.setCapStyle(Qt.RoundCap); pen.setJoinStyle(Qt.RoundJoin)
    p.setPen(pen); p.setBrush(Qt.NoBrush)
    path = QPainterPath()
    path.moveTo(s * 0.22, s * 0.38); path.lineTo(s * 0.50, s * 0.66)
    path.lineTo(s * 0.78, s * 0.38)
    p.drawPath(path)
    p.end(); return QIcon(pm)


def icon_upload(size=56, color=None):
    color = color or C_GREEN
    pm = QPixmap(size, size); pm.fill(Qt.transparent)
    p = QPainter(pm); p.setRenderHint(QPainter.Antialiasing)
    s = size
    pen = QPen(QColor(color), s * 0.065)
    pen.setCapStyle(Qt.RoundCap); pen.setJoinStyle(Qt.RoundJoin)
    p.setPen(pen); p.setBrush(Qt.NoBrush)
    p.drawLine(int(s * 0.50), int(s * 0.62), int(s * 0.50), int(s * 0.18))
    p.drawLine(int(s * 0.50), int(s * 0.18), int(s * 0.30), int(s * 0.38))
    p.drawLine(int(s * 0.50), int(s * 0.18), int(s * 0.70), int(s * 0.38))
    path = QPainterPath()
    path.moveTo(s * 0.20, s * 0.74); path.lineTo(s * 0.20, s * 0.84)
    path.lineTo(s * 0.80, s * 0.84); path.lineTo(s * 0.80, s * 0.74)
    p.drawPath(path)
    p.end(); return QIcon(pm)


def icon_arrow_left(size=18, color=None):
    color = color or C_DARK_GREEN
    pm = QPixmap(size, size); pm.fill(Qt.transparent)
    p = QPainter(pm); p.setRenderHint(QPainter.Antialiasing)
    s = size
    pen = QPen(QColor(color), s * 0.12)
    pen.setCapStyle(Qt.RoundCap); pen.setJoinStyle(Qt.RoundJoin)
    p.setPen(pen); p.setBrush(Qt.NoBrush)
    p.drawLine(int(s * 0.70), int(s * 0.50), int(s * 0.25), int(s * 0.50))
    p.drawLine(int(s * 0.25), int(s * 0.50), int(s * 0.45), int(s * 0.30))
    p.drawLine(int(s * 0.25), int(s * 0.50), int(s * 0.45), int(s * 0.70))
    p.end(); return QIcon(pm)


def icon_gear(size=22, color=None):
    color = color or C_DARK_GREEN
    pm = QPixmap(size, size); pm.fill(Qt.transparent)
    p = QPainter(pm); p.setRenderHint(QPainter.Antialiasing)
    s = size
    cx, cy = s * 0.5, s * 0.5
    r_outer = s * 0.42
    r_ring  = s * 0.28
    r_hole  = s * 0.11
    pen = QPen(QColor(color), s * 0.095)
    pen.setCapStyle(Qt.RoundCap); pen.setJoinStyle(Qt.RoundJoin)
    p.setPen(pen); p.setBrush(Qt.NoBrush)
    teeth = 8
    for i in range(teeth):
        a = 2 * math.pi * i / teeth - math.pi / 2
        x1 = cx + math.cos(a) * r_ring
        y1 = cy + math.sin(a) * r_ring
        x2 = cx + math.cos(a) * r_outer
        y2 = cy + math.sin(a) * r_outer
        p.drawLine(QPoint(int(x1), int(y1)), QPoint(int(x2), int(y2)))
    p.drawEllipse(QPoint(int(cx), int(cy)), int(r_ring), int(r_ring))
    p.drawEllipse(QPoint(int(cx), int(cy)), int(r_hole), int(r_hole))
    p.end(); return QIcon(pm)


def icon_status(size=42, kind="info"):
    T = get_theme()
    color = {
        "info":     T["primary"],
        "warning":  C_ORANGE,
        "error":    C_RED,
        "delete":   C_RED,
        "question": T["primary"],
    }.get(kind, T["primary"])

    pm = QPixmap(size, size); pm.fill(Qt.transparent)
    p = QPainter(pm); p.setRenderHint(QPainter.Antialiasing)
    s = size
    p.setPen(Qt.NoPen); p.setBrush(QBrush(QColor(color)))
    p.drawEllipse(0, 0, s, s)
    pen = QPen(QColor("#FFFFFF"), s * 0.10)
    pen.setCapStyle(Qt.RoundCap); pen.setJoinStyle(Qt.RoundJoin)
    p.setPen(pen); p.setBrush(Qt.NoBrush)
    if kind == "info":
        p.drawLine(int(s * 0.5), int(s * 0.26), int(s * 0.5), int(s * 0.30))
        p.drawLine(int(s * 0.5), int(s * 0.44), int(s * 0.5), int(s * 0.74))
    elif kind in ("warning", "question"):
        p.drawLine(int(s * 0.5), int(s * 0.26), int(s * 0.5), int(s * 0.58))
        p.drawLine(int(s * 0.5), int(s * 0.70), int(s * 0.5), int(s * 0.74))
    elif kind == "error":
        p.drawLine(int(s * 0.32), int(s * 0.32), int(s * 0.68), int(s * 0.68))
        p.drawLine(int(s * 0.68), int(s * 0.32), int(s * 0.32), int(s * 0.68))
    elif kind == "delete":
        p.drawLine(int(s * 0.30), int(s * 0.38), int(s * 0.70), int(s * 0.38))
        p.drawLine(int(s * 0.42), int(s * 0.38), int(s * 0.42), int(s * 0.30))
        p.drawLine(int(s * 0.58), int(s * 0.38), int(s * 0.58), int(s * 0.30))
        p.drawLine(int(s * 0.42), int(s * 0.30), int(s * 0.58), int(s * 0.30))
        path = QPainterPath()
        path.moveTo(s * 0.36, s * 0.38); path.lineTo(s * 0.40, s * 0.74)
        path.lineTo(s * 0.60, s * 0.74); path.lineTo(s * 0.64, s * 0.38)
        p.drawPath(path)
        p.drawLine(int(s * 0.46), int(s * 0.46), int(s * 0.46), int(s * 0.68))
        p.drawLine(int(s * 0.54), int(s * 0.46), int(s * 0.54), int(s * 0.68))
    p.end(); return QIcon(pm)


# ═════════════════════════════════════════════════════════════════
#  دیالوگ تم‌دار
# ═════════════════════════════════════════════════════════════════

class ThemedDialog(QDialog):
    def __init__(self, parent, title, message, kind="info",
                 buttons=(("تایید", True, QDialog.Accepted),)):
        super().__init__(parent)
        self.setWindowFlags(Qt.Dialog | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setModal(True)
        self.setLayoutDirection(Qt.RightToLeft)
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
        close_btn = QPushButton("✕")
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
                         buttons=(("بله", True, QDialog.Accepted),
                                  ("انصراف", False, QDialog.Rejected)))
        return d.exec() == QDialog.Accepted

    @staticmethod
    def confirm_delete(parent, title, message):
        d = ThemedDialog(parent, title, message, "delete",
                         buttons=(("حذف", True, QDialog.Accepted),
                                  ("انصراف", False, QDialog.Rejected)))
        return d.exec() == QDialog.Accepted

    def showEvent(self, e):
        super().showEvent(e)
        par = self.parent()
        if par:
            pg = par.window().geometry()
            self.move(pg.center().x() - self.width() // 2,
                      pg.center().y() - self.height() // 2)


# ═════════════════════════════════════════════════════════════════
#  مدیریت JSON
# ═════════════════════════════════════════════════════════════════

def load_templates_raw():
    if not TEMPLATES_JSON.exists():
        return []
    try:
        with open(TEMPLATES_JSON, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except Exception as e:
        print(f"[خطا] خواندن JSON: {e}")
        return []


def load_templates():
    return [t for t in load_templates_raw()
            if isinstance(t, dict) and "name" in t and "address" in t]


def save_templates(templates):
    WORDCLONER_DIR.mkdir(parents=True, exist_ok=True)
    with open(TEMPLATES_JSON, "w", encoding="utf-8") as f:
        json.dump(templates, f, ensure_ascii=False, indent=4)


# ═════════════════════════════════════════════════════════════════
#  پردازش Word
# ═════════════════════════════════════════════════════════════════

def normalize_quotes(text):
    return (text
            .replace("\u201c", '"').replace("\u201d", '"')
            .replace("\u201e", '"').replace("\u201f", '"')
            .replace("\u2033", '"'))


def _get_wrapper():
    return SETTINGS.get("placeholder_wrapper", '""')


def extract_placeholders(docx_path):
    doc = Document(docx_path)
    raw_text = "\n".join(p.text for p in doc.paragraphs)
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

    props = {}
    for i in range(0, len(positions) - 1, 2):
        open_pos = positions[i]
        close_pos = positions[i + 1]
        term_start = open_pos + wlen
        term = text[term_start:close_pos]
        if term and term not in props:
            props[term] = term_start
    return props


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
            key = norm[start + wlen:end]
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
            override = font_overrides.get(key, (None, False, False))
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
            override = font_overrides.get(key, (None, False, False))
            size_pt = auto_sizes[(start, end, key)]
            ref_run = auto_refs.get((start, end, key))
            value = replacements[key]
            segments.append(
                (value, (size_pt, override[1], override[2], ref_run), True)
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
                size_pt, bold, italic, ref_run = font_spec
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


# ═════════════════════════════════════════════════════════════════
#  نوار عنوان پنجره  (دکمه‌ها سمت راست)
# ═════════════════════════════════════════════════════════════════

class TitleBar(QFrame):
    def __init__(self, window):
        super().__init__()
        self._win = window; self._drag = None
        self.setFixedHeight(46); self.setObjectName("titleBar")

        lay = QHBoxLayout(self)
        lay.setContentsMargins(8, 0, 16, 0); lay.setSpacing(10)

        # در RTL: ویجتی که اول اضافه شود، سمت راست قرار می‌گیرد
        # ترتیب بصری از چپ به راست: [آیکون][عنوان][stretch][min][max][close]
        b_cls = self._btn("✕", C_RED)
        b_max = self._btn("☐", C_GREEN)
        b_min = self._btn("−", C_GREEN)
        b_min.clicked.connect(self._win.showMinimized)
        b_max.clicked.connect(self._toggle_max)
        b_cls.clicked.connect(self._win.close)

        # ابتدا دکمه‌ها (سمت راست)
        lay.addWidget(b_cls)
        lay.addWidget(b_max)
        lay.addWidget(b_min)
        lay.addStretch()

        # سپس آیکون و عنوان (سمت چپ)
        ico = QLabel()
        ico.setPixmap(icon_document(20, "#FFFFFF").pixmap(QSize(20, 20)))
        ico.setStyleSheet("background: transparent;")

        t = QLabel("سامانه صدور نامه")
        f = QFont(); f.setPointSize(11); f.setBold(True)
        t.setFont(f); t.setStyleSheet("color: white; background: transparent;")

        lay.addWidget(t)
        lay.addWidget(ico)

        self.setStyleSheet(f"""
            QFrame#titleBar {{
                background-color: {C_DARK_GREEN};
                border-top-left-radius: 14px;
                border-top-right-radius: 14px;
            }}
        """)

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
        if e.button() == Qt.LeftButton:
            self._drag = e.globalPosition().toPoint() - self._win.frameGeometry().topLeft()

    def mouseMoveEvent(self, e):
        if e.buttons() & Qt.LeftButton and self._drag and not self._win.isMaximized():
            self._win.move(e.globalPosition().toPoint() - self._drag)

    def mouseReleaseEvent(self, e):
        self._drag = None

    def mouseDoubleClickEvent(self, e):
        self._toggle_max()


# ═════════════════════════════════════════════════════════════════
#  دکمه‌ها و لیست کشویی
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

        self.text = QLabel("انتخاب قالب نامه")
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


class PopupItem(QFrame):
    clicked = Signal()
    delete_clicked = Signal()

    def __init__(self, icon: QIcon, text: str, *, is_add=False, show_delete=False):
        super().__init__()
        self.setObjectName("it")
        self.setFixedHeight(46)
        self.setCursor(Qt.PointingHandCursor)

        bg       = C_LIGHT_GREEN if is_add else "transparent"
        bg_hover = C_ACCENT1 if is_add else C_HOVER
        self.setStyleSheet(f"""
            QFrame#it {{ background-color: {bg}; border-radius: 8px; }}
            QFrame#it:hover {{ background-color: {bg_hover}; }}
        """)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(10, 0, 10, 0); lay.setSpacing(10)

        ic = QLabel(); ic.setPixmap(icon.pixmap(QSize(26, 26)))
        ic.setStyleSheet("background: transparent;")
        ic.setAttribute(Qt.WA_TransparentForMouseEvents)

        tx = QLabel(text)
        f = QFont(); f.setPointSize(11)
        if is_add: f.setBold(True)
        tx.setFont(f)
        tx.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")
        tx.setAttribute(Qt.WA_TransparentForMouseEvents)

        lay.addWidget(ic); lay.addWidget(tx); lay.addStretch()

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

    ADD_ITEM_H = 46
    SEP_H      = 10
    PADDING_H  = 16
    TEMPLATE_H = 48

    def __init__(self, templates, parent=None):
        super().__init__(
            parent,
            Qt.Popup | Qt.FramelessWindowHint | Qt.NoDropShadowWindowHint
        )
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setLayoutDirection(Qt.RightToLeft)

        self._templates = templates
        self._scroll = None
        self._content_h = 0

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

        add_item = PopupItem(icon_plus_circle(28), "افزودن نامه", is_add=True)
        add_item.clicked.connect(self.add_clicked.emit)
        inner.addWidget(add_item)

        if templates:
            sep = QFrame(); sep.setFixedHeight(1)
            sep.setStyleSheet(f"background-color: {C_BORDER}; margin: 4px 10px;")
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
                it = PopupItem(icon_document(22), tpl.get("name", "بدون نام"), show_delete=True)
                it.clicked.connect(lambda _=None, t=tpl: self.template_selected.emit(t))
                it.delete_clicked.connect(lambda _=None, t=tpl: self.template_deleted.emit(t))
                cl.addWidget(it)
            cl.addStretch()

            scroll.setWidget(content)
            self._scroll = scroll
            self._content_h = len(templates) * self.TEMPLATE_H + 8
            scroll.setFixedHeight(self._content_h)
            inner.addWidget(scroll)

    def constrain_height(self, available_h):
        if self._scroll is None:
            return
        overhead = self.ADD_ITEM_H + self.SEP_H + self.PADDING_H
        max_scroll = max(60, int(available_h) - overhead)
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
        self.button.set_text(tpl.get("name", "انتخاب قالب نامه"))
        self.template_selected.emit(tpl)

    def _on_delete(self, tpl):
        self._close(); self.template_deleted.emit(tpl)


# ═════════════════════════════════════════════════════════════════
#  ناحیه Drag & Drop
# ═════════════════════════════════════════════════════════════════

class DropZone(QFrame):
    file_selected = Signal(str)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setObjectName("dropzone")
        self._file_path = None
        self.setMinimumHeight(220)
        self._apply_style(active=False)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 30, 24, 30); lay.setSpacing(14)

        self.icon_lbl = QLabel()
        self.icon_lbl.setPixmap(icon_upload(56).pixmap(QSize(56, 56)))
        self.icon_lbl.setAlignment(Qt.AlignCenter)
        self.icon_lbl.setStyleSheet("background: transparent; border: none;")
        self.icon_lbl.setAttribute(Qt.WA_TransparentForMouseEvents)

        self.text_lbl = QLabel("فایل ورد را اینجا رها کنید\nیا کلیک کنید تا از سیستم انتخاب شود")
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
                self, "انتخاب فایل ورد", start_dir, "Word Document (*.docx)")
            if path:
                self._set_file(path)

    def _set_file(self, path):
        self._file_path = path
        self.text_lbl.setText(f"✓  {Path(path).name}")
        f = QFont(); f.setPointSize(11); f.setBold(True)
        self.text_lbl.setFont(f)
        self.text_lbl.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent; border: none;")
        self.file_selected.emit(path)

    def file_path(self):
        return self._file_path

    def reset(self):
        self._file_path = None
        self.text_lbl.setText("فایل ورد را اینجا رها کنید\nیا کلیک کنید تا از سیستم انتخاب شود")
        f = QFont(); f.setPointSize(11)
        self.text_lbl.setFont(f)
        self.text_lbl.setStyleSheet(f"color: {C_MUTED}; background: transparent; border: none;")
        self._apply_style(active=False)


# ═════════════════════════════════════════════════════════════════
#  کامپوننت‌های کوچک برای فرم متغیرها
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


class FontSizeCombo(QComboBox):
    def __init__(self):
        super().__init__()
        self.setEditable(True)
        self.setFixedHeight(32)
        self.setMinimumWidth(86)
        self.setCursor(Qt.PointingHandCursor)
        self.setLayoutDirection(Qt.RightToLeft)
        self.addItems(FONT_SIZES)
        self.setCurrentText("Auto")
        self.setToolTip("سایز فونت (Auto = هم‌اندازه‌ی متن اطراف)")

        arrow_uri = _svg_arrow_uri(C_DARK_GREEN)
        self.setStyleSheet(f"""
            QComboBox {{
                background-color: #FFFFFF;
                border: 1.5px solid {C_BORDER};
                border-radius: 8px;
                padding: 0 6px 0 26px;
                color: {C_DARK_GREEN};
                font-size: 11px;
            }}
            QComboBox:focus {{ border-color: {C_GREEN}; }}
            QComboBox::drop-down {{
                subcontrol-origin: padding;
                subcontrol-position: center left;
                width: 24px;
                border: none;
                background: transparent;
            }}
            QComboBox::down-arrow {{
                image: {arrow_uri};
                width: 12px;
                height: 12px;
            }}
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


# ═════════════════════════════════════════════════════════════════
#  صفحه اصلی
# ═════════════════════════════════════════════════════════════════

class HomePage(QWidget):
    settings_requested = Signal()

    def __init__(self):
        super().__init__()
        self.setStyleSheet("background: transparent;")

        outer = QVBoxLayout(self)
        outer.setContentsMargins(40, 18, 40, 30)
        outer.setSpacing(0)

        top_row = QHBoxLayout()
        gear_btn = QPushButton()
        gear_btn.setIcon(icon_gear(22))
        gear_btn.setIconSize(QSize(22, 22))
        gear_btn.setFixedSize(42, 42)
        gear_btn.setCursor(Qt.PointingHandCursor)
        gear_btn.setToolTip("تنظیمات")
        gear_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                border: 1px solid transparent;
                border-radius: 10px;
            }}
            QPushButton:hover {{
                background-color: {C_LIGHT_GREEN};
                border-color: {C_BORDER};
            }}
        """)
        gear_btn.clicked.connect(self.settings_requested.emit)
        top_row.addWidget(gear_btn)
        top_row.addStretch()
        outer.addLayout(top_row)

        outer.addStretch()

        card = QWidget(); card.setStyleSheet("background: transparent;")
        card.setFixedWidth(440)
        cl = QVBoxLayout(card)
        cl.setContentsMargins(0, 0, 0, 0); cl.setSpacing(8)

        big = QLabel()
        big.setPixmap(icon_document_big(88).pixmap(QSize(88, 88)))
        big.setAlignment(Qt.AlignCenter)
        big.setStyleSheet("background: transparent;")

        title = QLabel("صدور نامه")
        tf = QFont(); tf.setPointSize(22); tf.setBold(True)
        title.setFont(tf); title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")

        self.dropdown = TemplateDropdown()

        cl.addWidget(big, 0, Qt.AlignHCenter)
        cl.addWidget(title); cl.addSpacing(18); cl.addWidget(self.dropdown)

        center = QHBoxLayout()
        center.addStretch(); center.addWidget(card); center.addStretch()
        outer.addLayout(center); outer.addStretch()


# ═════════════════════════════════════════════════════════════════
#  صفحه افزودن نامه
# ═════════════════════════════════════════════════════════════════

class AddLetterPage(QWidget):
    back_requested = Signal()
    saved = Signal()

    def __init__(self):
        super().__init__()
        self.setStyleSheet("background: transparent;")

        lay = QVBoxLayout(self)
        lay.setContentsMargins(40, 22, 40, 30); lay.setSpacing(14)

        header = QHBoxLayout()
        title = QLabel("افزودن نامه جدید")
        tf = QFont(); tf.setPointSize(15); tf.setBold(True)
        title.setFont(tf)
        title.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")

        back_btn = QPushButton("  بازگشت")
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

        nf = QFont(); nf.setPointSize(10); nf.setBold(True)

        name_lbl = QLabel("نام نامه")
        name_lbl.setFont(nf)
        name_lbl.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")
        lay.addWidget(name_lbl)

        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("مثلاً: نامه قبولی مقاله")
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
        lay.addWidget(self.name_input); lay.addSpacing(4)

        drop_lbl = QLabel("فایل قالب ورد")
        drop_lbl.setFont(nf)
        drop_lbl.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")
        lay.addWidget(drop_lbl)

        self.drop_zone = DropZone()
        lay.addWidget(self.drop_zone, 1)

        self.confirm_btn = QPushButton("تایید و ذخیره")
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
        lay.addWidget(self.confirm_btn)

    def reset(self):
        self.name_input.clear(); self.drop_zone.reset()

    def _on_confirm(self):
        if not HAS_DOCX:
            ThemedDialog.error(self, "کتابخانه ناقص",
                               "python-docx نصب نیست.\n\npip install python-docx")
            return

        name = self.name_input.text().strip()
        file_path = self.drop_zone.file_path()

        if not name:
            ThemedDialog.warning(self, "خطا", "لطفاً نام نامه را وارد کنید.")
            self.name_input.setFocus(); return
        if not file_path:
            ThemedDialog.warning(self, "خطا", "لطفاً فایل ورد را انتخاب یا رها کنید.")
            return

        existing = load_templates_raw()
        if any(t.get("name") == name for t in existing):
            ThemedDialog.warning(self, "خطا", f"نام «{name}» قبلاً ثبت شده است.")
            return

        try:
            WORDS_DIR.mkdir(parents=True, exist_ok=True)
            src = Path(file_path); dst = WORDS_DIR / src.name
            counter = 1
            while dst.exists():
                dst = WORDS_DIR / f"{src.stem}_{counter}{src.suffix}"
                counter += 1
            shutil.copy2(str(src), str(dst))

            props = extract_placeholders(str(dst))

            new_obj = {"name": name, "address": str(dst), "Properties": props}
            existing.append(new_obj)
            save_templates(existing)

            if props:
                ThemedDialog.info(self, "ذخیره شد",
                                  f"نامه «{name}» با {len(props)} متغیر ذخیره شد.")
            else:
                ThemedDialog.info(
                    self, "ذخیره شد",
                    f"نامه «{name}» ذخیره شد.\n"
                    f"توجه: هیچ متغیری با الگوی «{_get_wrapper()}» پیدا نشد.")

            self.reset(); self.saved.emit()
        except Exception as ex:
            ThemedDialog.error(self, "خطا", f"مشکلی رخ داد:\n{ex}")
            print("[خطا]", ex)


# ═════════════════════════════════════════════════════════════════
#  صفحه ویرایش نامه
# ═════════════════════════════════════════════════════════════════

class EditLetterPage(QWidget):
    back_requested = Signal()

    def __init__(self):
        super().__init__()
        self.setStyleSheet("background: transparent;")
        self._template = None
        self._fields = {}

        lay = QVBoxLayout(self)
        lay.setContentsMargins(40, 22, 40, 30); lay.setSpacing(14)

        header = QHBoxLayout()
        self.title_lbl = QLabel("صدور نامه")
        tf = QFont(); tf.setPointSize(15); tf.setBold(True)
        self.title_lbl.setFont(tf)
        self.title_lbl.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")

        back_btn = QPushButton("  بازگشت")
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

        header.addWidget(self.title_lbl); header.addStretch(); header.addWidget(back_btn)
        lay.addLayout(header)

        nf = QFont(); nf.setPointSize(10); nf.setBold(True)

        fn_lbl = QLabel("نام فایل خروجی")
        fn_lbl.setFont(nf)
        fn_lbl.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")
        lay.addWidget(fn_lbl)

        self.filename_input = QLineEdit()
        self.filename_input.setPlaceholderText("مثلاً: نامه_قبولی_مقاله")
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
        lay.addWidget(self.filename_input)

        vars_lbl = QLabel("مقادیر متغیرها")
        vars_lbl.setFont(nf)
        vars_lbl.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")
        lay.addWidget(vars_lbl)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        self.scroll.setStyleSheet(f"""
            QScrollArea {{ background: transparent; border: none; }}
            QScrollBar:vertical {{
                background: transparent; width: 8px; margin: 0;
            }}
            QScrollBar::handle:vertical {{
                background: {C_BORDER}; border-radius: 4px; min-height: 30px;
            }}
            QScrollBar::handle:vertical:hover {{ background: {C_GREEN}; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0;
            }}
        """)

        self.fields_container = QWidget()
        self.fields_container.setStyleSheet("background: transparent;")
        self.fields_layout = QVBoxLayout(self.fields_container)
        self.fields_layout.setContentsMargins(0, 0, 8, 0)
        self.fields_layout.setSpacing(12)
        self.fields_layout.addStretch()

        self.scroll.setWidget(self.fields_container)
        lay.addWidget(self.scroll, 1)

        btn_row = QHBoxLayout()
        self.confirm_btn = QPushButton("تایید و صدور نامه")
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
        btn_row.addWidget(self.confirm_btn)   # در RTL = سمت راست
        btn_row.addStretch()
        lay.addLayout(btn_row)

    def load_template(self, tpl):
        self._template = tpl
        self.title_lbl.setText(f"صدور: {tpl.get('name', '')}")
        self.filename_input.setText(tpl.get("name", ""))

        self._fields.clear()
        while self.fields_layout.count() > 1:
            item = self.fields_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

        props = tpl.get("Properties", {})
        if not props:
            empty = QLabel("این قالب متغیری ندارد.")
            empty.setStyleSheet(f"color: {C_MUTED}; background: transparent; padding: 8px;")
            empty.setFont(QFont("", 10))
            self.fields_layout.insertWidget(0, empty)
            return

        sorted_items = sorted(props.items(), key=lambda kv: kv[1])
        for idx, (key, _start) in enumerate(sorted_items):
            box = self._build_field(key, show_labels=(idx == 0))
            self.fields_layout.insertWidget(self.fields_layout.count() - 1, box)

    def _build_field(self, key, show_labels=False):
        box = QWidget()
        box.setStyleSheet("background: transparent;")
        bl = QVBoxLayout(box)
        bl.setContentsMargins(0, 0, 0, 0); bl.setSpacing(6)

        top = QHBoxLayout()
        top.setSpacing(8)

        lbl = QLabel(key)
        lf = QFont(); lf.setPointSize(10); lf.setBold(True)
        lbl.setFont(lf)
        lbl.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")
        top.addWidget(lbl)
        top.addStretch()

        # ─ برچسب «سایز فونت» + کمبو (فقط فیلد اول) ─
        if show_labels:
            size_lbl = QLabel("سایز فونت")
            size_lbl.setStyleSheet(
                f"color: {C_MUTED}; background: transparent; font-size: 10px;")
            top.addWidget(size_lbl)

        size_combo = FontSizeCombo()
        top.addWidget(size_combo)

        # ─ گروه بولد ─
        bold_wrap = QWidget()
        bold_wrap.setStyleSheet("background: transparent;")
        bold_lay = QHBoxLayout(bold_wrap)
        bold_lay.setContentsMargins(0, 0, 0, 0); bold_lay.setSpacing(6)
        bold_btn = RoundToggle("B", "بولد", bold=True)
        bold_lay.addWidget(bold_btn)
        if show_labels:
            bold_lbl = QLabel("بولد")
            bold_lbl.setStyleSheet(
                f"color: {C_MUTED}; background: transparent; font-size: 10px;")
            bold_lbl.setCursor(Qt.PointingHandCursor)
            bold_lbl.mousePressEvent = lambda _e, b=bold_btn: b.toggle()
            bold_lay.addWidget(bold_lbl)
        top.addWidget(bold_wrap)

        # ─ گروه ایتالیک ─
        ital_wrap = QWidget()
        ital_wrap.setStyleSheet("background: transparent;")
        ital_lay = QHBoxLayout(ital_wrap)
        ital_lay.setContentsMargins(0, 0, 0, 0); ital_lay.setSpacing(6)
        italic_btn = RoundToggle("I", "ایتالیک", italic=True)
        ital_lay.addWidget(italic_btn)
        if show_labels:
            italic_lbl = QLabel("ایتالیک")
            italic_lbl.setStyleSheet(
                f"color: {C_MUTED}; background: transparent; font-size: 10px;")
            italic_lbl.setCursor(Qt.PointingHandCursor)
            italic_lbl.mousePressEvent = lambda _e, b=italic_btn: b.toggle()
            ital_lay.addWidget(italic_lbl)
        top.addWidget(ital_wrap)

        bl.addLayout(top)

        inp = QLineEdit()
        inp.setPlaceholderText(f"مقدار «{key}»")
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
        }
        return box

    def _on_confirm(self):
        if not HAS_DOCX:
            ThemedDialog.error(self, "کتابخانه ناقص",
                               "python-docx نصب نیست.\n\npip install python-docx")
            return
        if not self._template:
            return

        filename = self.filename_input.text().strip()
        if not filename:
            ThemedDialog.warning(self, "خطا", "لطفاً نام فایل خروجی را وارد کنید.")
            return

        # ─ الزام پر بودن حداقل یک فیلد متغیر ─
        if self._fields:
            any_filled = any(
                w["input"].text().strip() for w in self._fields.values()
            )
            if not any_filled:
                ThemedDialog.warning(
                    self, "خطا",
                    "لطفاً مقدار حداقل یکی از متغیرها را وارد کنید.")
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
            ThemedDialog.error(self, "خطا", f"ساخت مسیر خروجی ناموفق:\n{e}")
            return

        out_path = output_dir / filename
        src = Path(self._template.get("address", ""))
        if not src.exists():
            ThemedDialog.error(self, "خطا", f"فایل قالب پیدا نشد:\n{src}")
            return

        if out_path.exists():
            ok = ThemedDialog.confirm(
                self, "فایل موجود است",
                f"فایل «{filename}» در مسیر خروجی وجود دارد.\nآن را بازنویسی کنیم؟")
            if not ok: return

        replacements = {}
        font_overrides = {}
        for key, widgets in self._fields.items():
            replacements[key] = widgets["input"].text()

            size_text = widgets["size"].currentText().strip().lower()
            if size_text in ("", "auto"):
                size_val = None
            else:
                try:
                    size_val = float(size_text.replace(",", "."))
                except ValueError:
                    size_val = None
            bold_val = widgets["bold"].isChecked()
            italic_val = widgets["italic"].isChecked()
            font_overrides[key] = (size_val, bold_val, italic_val)

        try:
            replace_placeholders_in_docx(
                str(src), str(out_path), replacements, font_overrides)
            ThemedDialog.info(
                self, "صادر شد",
                f"نامه با موفقیت ذخیره شد:\n{out_path}")
            self.back_requested.emit()
        except Exception as ex:
            ThemedDialog.error(self, "خطا", f"مشکلی در صدور نامه رخ داد:\n{ex}")
            print("[خطا]", ex)


# ═════════════════════════════════════════════════════════════════
#  صفحه تنظیمات
# ═════════════════════════════════════════════════════════════════

class SettingsPage(QWidget):
    back_requested = Signal()
    saved = Signal()

    IN_H      = 36
    THEME_W   = 220
    OUTPUT_W  = 380
    WRAPPER_W = 180

    def __init__(self):
        super().__init__()
        self.setStyleSheet("background: transparent;")

        lay = QVBoxLayout(self)
        lay.setContentsMargins(40, 22, 40, 30); lay.setSpacing(10)

        header = QHBoxLayout()
        title = QLabel("تنظیمات")
        tf = QFont(); tf.setPointSize(15); tf.setBold(True)
        title.setFont(tf)
        title.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")

        back_btn = QPushButton("  بازگشت")
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

        # ─ تم ─
        theme_lbl = QLabel("تم برنامه")
        theme_lbl.setFont(lf)
        theme_lbl.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")
        lay.addWidget(theme_lbl)

        self.theme_combo = QComboBox()
        self.theme_combo.setFixedHeight(self.IN_H)
        self.theme_combo.setFixedWidth(self.THEME_W)
        self.theme_combo.setCursor(Qt.PointingHandCursor)
        self.theme_combo.setLayoutDirection(Qt.RightToLeft)
        for key, t in THEMES.items():
            self.theme_combo.addItem(t["label"], key)
        cur = SETTINGS.get("theme", "green")
        idx = self.theme_combo.findData(cur)
        if idx >= 0:
            self.theme_combo.setCurrentIndex(idx)

        arrow_uri = _svg_arrow_uri(C_DARK_GREEN)
        self.theme_combo.setStyleSheet(f"""
            QComboBox {{
                background-color: #FFFFFF;
                border: 1.5px solid {C_BORDER};
                border-radius: 8px;
                padding: 0 6px 0 26px;
                color: {C_DARK_GREEN};
                font-size: 11px;
            }}
            QComboBox:focus, QComboBox:on {{ border-color: {C_GREEN}; }}
            QComboBox::drop-down {{
                subcontrol-origin: padding;
                subcontrol-position: center left;
                width: 24px;
                border: none;
                background: transparent;
            }}
            QComboBox::down-arrow {{
                image: {arrow_uri};
                width: 12px;
                height: 12px;
            }}
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

        theme_row = QHBoxLayout()
        theme_row.setContentsMargins(0, 0, 0, 0)
        theme_row.addWidget(self.theme_combo)
        theme_row.addStretch()
        lay.addLayout(theme_row); lay.addSpacing(2)

        # ─ مسیر خروجی ─
        out_lbl = QLabel("مسیر خروجی نامه‌ها")
        out_lbl.setFont(lf)
        out_lbl.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")
        lay.addWidget(out_lbl)

        out_row = QHBoxLayout(); out_row.setSpacing(8)
        out_row.setContentsMargins(0, 0, 0, 0)

        self.output_input = QLineEdit()
        self.output_input.setText(SETTINGS.get("output_dir", str(Path.home() / "Desktop")))
        self.output_input.setFixedHeight(self.IN_H)
        self.output_input.setFixedWidth(self.OUTPUT_W)
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
        out_row.addWidget(self.output_input)

        browse_btn = QPushButton("انتخاب…")
        browse_btn.setFixedHeight(self.IN_H)
        browse_btn.setFixedWidth(80)
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
        out_row.addStretch()
        lay.addLayout(out_row); lay.addSpacing(2)

        # ─ عبارت دربرگیرنده ─
        wrap_lbl = QLabel("عبارت دربرگیرنده متغیرها")
        wrap_lbl.setFont(lf)
        wrap_lbl.setStyleSheet(f"color: {C_DARK_GREEN}; background: transparent;")
        lay.addWidget(wrap_lbl)

        self.wrapper_input = QLineEdit()
        self.wrapper_input.setText(SETTINGS.get("placeholder_wrapper", '""'))
        self.wrapper_input.setFixedHeight(self.IN_H)
        self.wrapper_input.setFixedWidth(self.WRAPPER_W)
        self.wrapper_input.setPlaceholderText('مثلاً: ""')
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

        hint = QLabel(
            "متغیرها در فایل Word بین این عبارت نوشته می‌شوند.\n"
            "مثال:  " + self.wrapper_input.text() + "نام کاربر" + self.wrapper_input.text()
        )
        hint.setStyleSheet(f"color: {C_MUTED}; background: transparent; font-size: 10px;")
        hint.setWordWrap(True)
        lay.addWidget(hint)

        lay.addStretch()

        # ─ دکمه ذخیره: در RTL با addWidget اول، سمت راست قرار می‌گیرد ─
        btn_row = QHBoxLayout()
        self.save_btn = QPushButton("ذخیره تنظیمات")
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
        btn_row.addWidget(self.save_btn)   # سمت راست در RTL
        btn_row.addStretch()
        lay.addLayout(btn_row)

    def _pick_output_dir(self):
        start = self.output_input.text() or str(Path.home())
        d = QFileDialog.getExistingDirectory(self, "انتخاب مسیر خروجی", start)
        if d:
            self.output_input.setText(d)

    def _on_save(self):
        new_theme = self.theme_combo.currentData()
        new_output = self.output_input.text().strip()
        new_wrapper = self.wrapper_input.text()

        if not new_output:
            ThemedDialog.warning(self, "خطا", "مسیر خروجی نمی‌تواند خالی باشد.")
            return
        try:
            Path(new_output).mkdir(parents=True, exist_ok=True)
        except Exception as e:
            ThemedDialog.error(self, "خطا", f"مسیر خروجی قابل استفاده نیست:\n{e}")
            return

        if len(new_wrapper) == 0:
            ThemedDialog.warning(self, "خطا", "عبارت دربرگیرنده نمی‌تواند خالی باشد.")
            return

        SETTINGS["theme"] = new_theme
        SETTINGS["output_dir"] = new_output
        SETTINGS["placeholder_wrapper"] = new_wrapper
        save_settings()

        self.saved.emit()


# ═════════════════════════════════════════════════════════════════
#  ظرف اصلی
# ═════════════════════════════════════════════════════════════════

class Container(QWidget):
    def __init__(self):
        super().__init__()
        self.setStyleSheet("background: transparent;")

    def paintEvent(self, e):
        p = QPainter(self); p.setRenderHint(QPainter.Antialiasing)
        w, h = self.width(), self.height()

        path = QPainterPath()
        path.addRoundedRect(QRectF(0.5, 0.5, w - 1, h - 1), 14, 14)
        p.fillPath(path, QColor(C_BG))
        p.setClipPath(path)

        c1 = QPainterPath()
        c1.moveTo(0, h); c1.lineTo(0, h * 0.55)
        c1.cubicTo(w * 0.08, h * 0.72, w * 0.20, h * 0.85, w * 0.34, h)
        c1.closeSubpath()
        p.fillPath(c1, QColor(C_DARK_GREEN))

        c2 = QPainterPath()
        c2.moveTo(0, h); c2.lineTo(0, h * 0.72)
        c2.cubicTo(w * 0.06, h * 0.83, w * 0.16, h * 0.92, w * 0.24, h)
        c2.closeSubpath()
        p.fillPath(c2, QColor(C_ACCENT1))

        c3 = QPainterPath()
        c3.moveTo(w, 0); c3.lineTo(w * 0.90, 0)
        c3.cubicTo(w * 0.92, h * 0.06, w * 0.96, h * 0.12, w, h * 0.18)
        c3.closeSubpath()
        p.fillPath(c3, QColor(C_ACCENT2))

        p.setClipping(False)
        p.setPen(QPen(QColor(C_BORDER), 1)); p.setBrush(Qt.NoBrush)
        p.drawPath(path)


# ═════════════════════════════════════════════════════════════════
#  پنجره اصلی
# ═════════════════════════════════════════════════════════════════

class MainWindow(QMainWindow):
    theme_changed = Signal()

    def __init__(self):
        super().__init__()
        self.setWindowTitle("سامانه صدور نامه")
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.resize(1100, 720)
        self.setMinimumSize(820, 580)

        root = QWidget()
        root.setAttribute(Qt.WA_TranslucentBackground)
        self._root_layout = QVBoxLayout(root)
        self._root_layout.setContentsMargins(16, 16, 16, 16)

        self.container = Container()
        sh = QGraphicsDropShadowEffect(self)
        sh.setBlurRadius(40); sh.setOffset(0, 10)
        sh.setColor(QColor(0, 0, 0, 70))
        self.container.setGraphicsEffect(sh)
        self._root_layout.addWidget(self.container)
        self.setCentralWidget(root)

        v = QVBoxLayout(self.container)
        v.setContentsMargins(0, 0, 0, 0); v.setSpacing(0)
        v.addWidget(TitleBar(self))

        self.stack = QStackedWidget()
        self.stack.setStyleSheet("background: transparent;")
        v.addWidget(self.stack, 1)

        self.home_page = HomePage()
        self.add_page = AddLetterPage()
        self.edit_page = EditLetterPage()
        self.settings_page = SettingsPage()

        for w in (self.home_page, self.add_page, self.edit_page, self.settings_page):
            self.stack.addWidget(w)

        self.home_page.dropdown.add_requested.connect(self._go_to_add)
        self.home_page.dropdown.template_selected.connect(self._on_template_selected)
        self.home_page.dropdown.template_deleted.connect(self._on_template_deleted)
        self.home_page.settings_requested.connect(self._go_to_settings)

        self.add_page.back_requested.connect(self._go_to_home)
        self.add_page.saved.connect(self._on_letter_saved)

        self.edit_page.back_requested.connect(self._go_to_home)

        self.settings_page.back_requested.connect(self._go_to_home)
        self.settings_page.saved.connect(self._on_settings_saved)

        self.reload_templates()

    def reload_templates(self):
        self.home_page.dropdown.set_templates(load_templates())

    def _go_to_home(self):
        self.stack.setCurrentWidget(self.home_page)

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
        current_dark = C_DARK_GREEN
        new_dark = THEMES.get(SETTINGS.get("theme"), {}).get("dark")
        if new_dark and new_dark != current_dark:
            self.theme_changed.emit()
        else:
            self._go_to_home()

    def _on_letter_saved(self):
        self.reload_templates()
        self._go_to_home()

    def _on_template_selected(self, tpl):
        if not HAS_DOCX:
            ThemedDialog.error(self, "کتابخانه ناقص",
                               "python-docx نصب نیست.\n\npip install python-docx")
            return
        self.edit_page.load_template(tpl)
        self.stack.setCurrentWidget(self.edit_page)

    def _on_template_deleted(self, tpl):
        name = tpl.get("name", "")
        address = tpl.get("address", "")

        ok = ThemedDialog.confirm_delete(
            self, "حذف قالب",
            f"آیا از حذف «{name}» مطمئن هستید؟\n"
            "فایل ورد مربوطه نیز از پوشه Words حذف خواهد شد."
        )
        if not ok: return

        data = load_templates_raw()
        others = [t for t in data
                  if t.get("name") != name and t.get("address") == address]

        data = [t for t in data if t.get("name") != name]
        save_templates(data)

        file_deleted = False
        if not others and address:
            p = Path(address)
            if p.exists():
                try:
                    p.unlink()
                    file_deleted = True
                except Exception as e:
                    print(f"[خطا حذف فایل] {e}")

        self.reload_templates()

        msg = f"قالب «{name}» با موفقیت حذف شد."
        if file_deleted:
            msg += "\nفایل ورد نیز از پوشه Words پاک شد."
        elif others:
            msg += "\n(فایل ورد حفظ شد چون قالب دیگری از آن استفاده می‌کند.)"
        ThemedDialog.info(self, "حذف شد", msg)

    def changeEvent(self, e):
        if e.type() == QEvent.WindowStateChange:
            m = 0 if self.isMaximized() else 16
            self._root_layout.setContentsMargins(m, m, m, m)
        super().changeEvent(e)


# ═════════════════════════════════════════════════════════════════
#  اجرا
# ═════════════════════════════════════════════════════════════════

def main():
    load_settings()
    reload_theme_colors()

    app = QApplication(sys.argv)
    app.setLayoutDirection(Qt.RightToLeft)
    app.setQuitOnLastWindowClosed(True)

    for fam in ("Vazirmatn", "IRANSans", "Tahoma", "Segoe UI"):
        f = QFont(fam, 10)
        if f.exactMatch() or fam == "Segoe UI":
            app.setFont(f)
            break

    current = [None]

    def restart():
        old = current[0]
        geo = old.geometry()
        was_max = old.isMaximized()
        old.hide()
        old.deleteLater()

        load_settings()
        reload_theme_colors()

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
        win.show()
        return win

    current[0] = create_window()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()