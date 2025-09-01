"""Application resources helpers."""

from __future__ import annotations

import os
import logging
from typing import Dict

from PySide6 import QtGui
from PySide6.QtGui import QIcon, QFont, QGuiApplication

logger = logging.getLogger(__name__)

ASSETS_DIR = os.path.join(os.path.dirname(__file__), "..", "assets")
FONTS_DIR = os.path.join(ASSETS_DIR, "fonts")
EXO2_FONTS_DIR = os.path.join(FONTS_DIR, "Exo2")
ICONS_DIR = os.path.join(ASSETS_DIR, "icons")

ICONS: Dict[str, QIcon] = {}


def register_fonts() -> None:
    """Register the Exo2 regular font and set a fallback on failure."""
    if not os.path.isdir(EXO2_FONTS_DIR):
        return

    font_path = os.path.join(EXO2_FONTS_DIR, "Exo2-Regular.ttf")
    if not os.path.isfile(font_path):
        logger.error("Font file 'Exo2-Regular.ttf' not found at '%s'", font_path)
        QGuiApplication.setFont(QFont("Arial"))
        return

    fid = QtGui.QFontDatabase.addApplicationFont(font_path)
    if fid == -1:
        logger.error("Failed to load font 'Exo2-Regular.ttf' from '%s'", font_path)
        QGuiApplication.setFont(QFont("Arial"))

    if "Exo 2" not in QtGui.QFontDatabase.families():
        logger.error("Font 'Exo 2' not registered")
        QGuiApplication.setFont(QFont("Arial"))


def load_icons(theme: str = "dark") -> None:
    """Load themed icons into the global :data:`ICONS` dictionary."""
    theme_dir = os.path.join(ICONS_DIR, theme)
    if not os.path.isdir(theme_dir):
        return
    ICONS.clear()
    for name in [
        "settings",
        "chevron-left",
        "chevron-right",
        "chevron-up",
        "chevron-down",
        "save",
        "x",
    ]:
        for ext in ("svg", "png"):
            path = os.path.join(theme_dir, f"{name}.{ext}")
            if os.path.exists(path):
                ICONS[name] = QIcon(path)
                break


def icon(name: str) -> QIcon:
    """Return previously loaded icon by *name*."""
    return ICONS.get(name, QIcon())
