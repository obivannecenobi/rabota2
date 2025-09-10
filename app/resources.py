"""Application resources helpers."""

from __future__ import annotations

import json
import os
import logging
from typing import Dict

from PySide6 import QtGui
from PySide6.QtGui import QIcon, QFont, QGuiApplication

logger = logging.getLogger(__name__)

ASSETS_DIR = os.path.join(os.path.dirname(__file__), "..", "assets")
FONTS_DIR = os.path.join(ASSETS_DIR, "fonts")
ICONS_DIR = os.path.join(ASSETS_DIR, "icons")

ICONS: Dict[str, QIcon] = {}


def register_fonts() -> None:
    """Register bundled fonts and ensure the default family is available.

    Recursively walks through :data:`FONTS_DIR` and loads every font file with a
    supported extension.  If the default "Exo 2" family cannot be registered,
    the application falls back to ``"Arial"`` and updates the global
    configuration accordingly.
    """

    def _set_fallback() -> None:
        """Apply the Arial fallback and store it in global CONFIG."""
        QGuiApplication.setFont(QFont("Arial"))
        try:  # deferred import to avoid circular dependency
            from . import main as _main

            _main.CONFIG["font_family"] = "Arial"
        except Exception:  # pragma: no cover - extremely defensive
            pass

    if not os.path.isdir(FONTS_DIR):
        _set_fallback()
        return

    extensions = {".ttf", ".otf", ".fon", ".ttc"}
    families: set[str] = set()
    for root, _dirs, files in os.walk(FONTS_DIR):
        for name in files:
            if os.path.splitext(name)[1].lower() not in extensions:
                continue
            path = os.path.join(root, name)
            fid = QtGui.QFontDatabase.addApplicationFont(path)
            if fid == -1:
                logger.error("Failed to load font '%s'", path)
                continue
            fams = QtGui.QFontDatabase.applicationFontFamilies(fid)
            families.update(fams)
            if name == "Cattedrale[RUSbypenka220]-Regular.ttf" and fams:
                try:  # deferred import to avoid circular dependency
                    from . import main as _main

                    _main.CONFIG["header_font"] = fams[0]
                    _main.CONFIG["sidebar_font"] = fams[0]
                    with open(_main.CONFIG_PATH, "r", encoding="utf-8") as fh:
                        cfg = json.load(fh)
                    changed = False
                    if cfg.get("header_font") != fams[0]:
                        cfg["header_font"] = fams[0]
                        changed = True
                    if cfg.get("sidebar_font") != fams[0]:
                        cfg["sidebar_font"] = fams[0]
                        changed = True
                    if changed:
                        with open(_main.CONFIG_PATH, "w", encoding="utf-8") as fh:
                            json.dump(cfg, fh, ensure_ascii=False, indent=2)
                except Exception:  # pragma: no cover - extremely defensive
                    pass

    if "Exo 2" not in families:
        logger.error("Font 'Exo 2' not registered")
        _set_fallback()
        return

    # Success – ensure CONFIG reflects the available family
    try:  # pragma: no cover - defensive
        from . import main as _main

        _main.CONFIG.setdefault("font_family", "Exo 2")
    except Exception:
        pass


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
