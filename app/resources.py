"""Application resources helpers."""

from __future__ import annotations

import os
import logging
import ctypes
from typing import Dict

logger = logging.getLogger(__name__)

try:  # pragma: no cover - optional dependency
    import tkinter as tk
    from tkinter import font as tkfont
except Exception as exc:  # pragma: no cover - extremely defensive
    tk = None  # type: ignore[assignment]
    tkfont = None  # type: ignore[assignment]
    logger.error("Failed to import tkinter: %s", exc)

try:  # pragma: no cover - optional dependency
    import customtkinter as ctk
except Exception as exc:  # pragma: no cover - extremely defensive
    ctk = None  # type: ignore[assignment]
    logger.error("Failed to import customtkinter: %s", exc)

from PySide6 import QtGui
from PySide6.QtGui import QIcon, QFont, QGuiApplication

ASSETS_DIR = os.path.join(os.path.dirname(__file__), "..", "assets")
FONTS_DIR = os.path.join(ASSETS_DIR, "fonts")
ICONS_DIR = os.path.join(ASSETS_DIR, "icons")

ICONS: Dict[str, QIcon] = {}


def register_cattedrale(font_path: str) -> str:
    if tk is None or ctk is None or tkfont is None:
        logger.error(
            "Skipping custom font registration for '%s': tkinter or customtkinter not available",
            font_path,
        )
        return "Exo 2"

    root = None
    try:
        if os.name == "nt":
            ctypes.windll.gdi32.AddFontResourceExW(font_path, 0x10, 0)
            ctypes.windll.user32.SendMessageW(0xFFFF, 0x1D, 0, 0)

        root = tk.Tk()
        root.withdraw()

        before = set(tkfont.families(root))
        ctk.FontManager.load_font(font_path)
        after = set(tkfont.families(root))
        new_fams = after - before
        family = next(iter(new_fams), os.path.splitext(os.path.basename(font_path))[0])
        return family
    except Exception:
        logger.exception("Failed to register font '%s'", font_path)
        return "Exo 2"
    finally:
        if root is not None:
            root.destroy()


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

    font_path = os.path.join(FONTS_DIR, "Cattedrale[RUSbypenka220]-Regular.ttf")
    if tk is not None and ctk is not None and tkfont is not None:
        fam = register_cattedrale(font_path)
        families.add(fam)
    else:
        logger.error(
            "Skipping custom font registration; tkinter or customtkinter not available"
        )

    for root, _dirs, files in os.walk(FONTS_DIR):
        for name in files:
            ext = os.path.splitext(name)[1].lower()
            if ext not in extensions:
                continue
            path = os.path.join(root, name)
            old_fams = set(QtGui.QFontDatabase.families())
            new_fams: set[str] = set()
            fid = QtGui.QFontDatabase.addApplicationFont(path) if ext != ".fon" else -1

            if ext == ".fon" or fid == -1:
                if os.name == "nt":
                    try:
                        ctypes.windll.gdi32.AddFontResourceExW(path, 0x10, 0)
                        ctypes.windll.user32.SendMessageW(0xFFFF, 0x1D, 0, 0)
                        new_fams = set(QtGui.QFontDatabase.families()) - old_fams
                    except Exception:  # pragma: no cover - extremely defensive
                        logger.error("Failed to load font '%s'", path)
                        continue
                else:
                    logger.error("Failed to load font '%s'", path)
                    continue
            else:
                new_fams = set(QtGui.QFontDatabase.applicationFontFamilies(fid))

            families.update(new_fams)

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
