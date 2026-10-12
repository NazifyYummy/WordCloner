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
