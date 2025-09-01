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
    """Register all Exo2 fonts found under :data:`EXO2_FONTS_DIR` recursively."""
    if not os.path.isdir(EXO2_FONTS_DIR):
        return

    fallback_set = False
    for root, _dirs, files in os.walk(EXO2_FONTS_DIR):
        for fname in files:
            if fname.lower().endswith((".ttf", ".otf")):
                path = os.path.join(root, fname)
                fid = QtGui.QFontDatabase.addApplicationFont(path)
                if fid == -1:
                    logger.error("Failed to load font '%s' from '%s'", fname, path)
                    if not fallback_set:
                        QGuiApplication.setFont(QFont("Sans Serif"))
                        fallback_set = True

    if "Exo 2" not in QtGui.QFontDatabase.families():
        logger.error("Font 'Exo 2' not registered")
        if not fallback_set:
            QGuiApplication.setFont(QFont("Sans Serif"))


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
