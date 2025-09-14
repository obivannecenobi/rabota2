"""Application resources helpers."""

from __future__ import annotations

import os
import logging
import ctypes
from typing import Dict
from concurrent.futures import ThreadPoolExecutor, TimeoutError

logger = logging.getLogger(__name__)

try:  # pragma: no cover - optional dependency
    import tkinter as tk
    from tkinter import font as tkfont
except Exception as exc:  # pragma: no cover - extremely defensive
    tk = None  # type: ignore[assignment]
    tkfont = None  # type: ignore[assignment]
    logger.warning("Failed to import tkinter: %s", exc)

try:  # pragma: no cover - optional dependency
    import customtkinter as ctk
except Exception as exc:  # pragma: no cover - extremely defensive
    ctk = None  # type: ignore[assignment]
    logger.warning("Failed to import customtkinter: %s", exc)

from PySide6 import QtGui, QtWidgets
from PySide6.QtGui import QIcon, QFont, QGuiApplication

ASSETS_DIR = os.path.join(os.path.dirname(__file__), "..", "assets")
FONTS_DIR = os.path.join(ASSETS_DIR, "fonts")
ICONS_DIR = os.path.join(ASSETS_DIR, "icons")

ICONS: Dict[str, QIcon] = {}


def _show_error_dialog(message: str) -> None:
    """Display a critical Qt dialog if a window is visible."""
    if os.name != "nt" and not os.environ.get("DISPLAY"):
        return
    app = QtWidgets.QApplication.instance()
    if not app:
        return
    for w in app.topLevelWidgets():
        if w.isVisible():
            QtWidgets.QMessageBox.critical(w, "Ошибка", message)
            break


def register_cattedrale(font_path: str) -> str:
    """Register the Cattedrale font for Tk-based widgets.

    Returns the detected family name or ``"Exo 2"`` if registration fails."""

    if tk is None or ctk is None or tkfont is None:
        logger.warning(
            "Не удалось загрузить tkinter/customtkinter — шрифт Cattedrale не будет применён"
        )
        return "Exo 2"

    if os.name != "nt" and not os.environ.get("DISPLAY"):
        logger.warning(
            "Переменная окружения DISPLAY не установлена — шрифт Cattedrale не будет применён"
        )
        return "Exo 2"

    def _worker() -> str:
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
            family = next(
                iter(new_fams), os.path.splitext(os.path.basename(font_path))[0]
            )
            return family
        finally:
            if root is not None:
                root.destroy()

    try:
        with ThreadPoolExecutor(max_workers=1) as executor:
            future = executor.submit(_worker)
            return future.result(timeout=5)
    except (Exception, TimeoutError):
        logger.exception("Failed to register font '%s'", font_path)
        return "Exo 2"


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
    if tk and ctk and tkfont and (os.name == "nt" or os.environ.get("DISPLAY")):
        fam = register_cattedrale(font_path)
    else:
        logger.warning(
            "Tkinter unavailable or running headless; skipping Cattedrale registration"
        )
        fam = "Exo 2"
    if fam == "Exo 2":
        logger.warning("Cattedrale font not found, using fallback")
    qt_family: str | None = None
    if fam != "Exo 2":
        fid = QtGui.QFontDatabase.addApplicationFont(font_path)
        if fid != -1:
            fams = QtGui.QFontDatabase.applicationFontFamilies(fid)
            if fams:
                qt_family = fams[0]
                families.add(qt_family)
    if qt_family is None and fam != "Exo 2":
        families.add(fam)

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

    # Success – ensure CONFIG reflects the available family and update header/sidebar
    try:  # pragma: no cover - defensive
        from . import main as _main
        from . import theme_manager

        _main.CONFIG.setdefault("font_family", "Exo 2")
        target_family = qt_family if qt_family is not None else fam
        if target_family == "Exo 2":
            _main.CONFIG["header_font"] = "Exo 2"
            _main.CONFIG["sidebar_font"] = "Exo 2"
        else:
            changed = False
            if _main.CONFIG.get("header_font") != target_family:
                _main.CONFIG["header_font"] = target_family
                changed = True
            if _main.CONFIG.get("sidebar_font") != target_family:
                _main.CONFIG["sidebar_font"] = target_family
                changed = True
            if changed:
                app = QtWidgets.QApplication.instance()
                if app:
                    for w in app.topLevelWidgets():
                        if hasattr(w, "apply_fonts"):
                            w.apply_fonts()

        theme_manager.set_header_font(_main.CONFIG["header_font"])
        theme_manager.set_text_font(_main.CONFIG["font_family"])
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
