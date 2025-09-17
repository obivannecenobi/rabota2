"""Application resources helpers."""

from __future__ import annotations

import os
import logging
import ctypes
from typing import Dict, Iterable
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

LATIN_RANGE = tuple(range(ord("A"), ord("Z") + 1)) + tuple(
    range(ord("a"), ord("z") + 1)
)
CYRILLIC_RANGE = (
    tuple(range(ord("А"), ord("Я") + 1))
    + tuple(range(ord("а"), ord("я") + 1))
    + (ord("Ё"), ord("ё"))
)
REQUIRED_RANGES = {
    "латинский": LATIN_RANGE,
    "кириллический": CYRILLIC_RANGE,
}


def _font_has_required_glyphs(family: str) -> tuple[bool, str | None]:
    """Return whether *family* provides required glyph ranges."""

    font = QFont(family)
    for label, codes in REQUIRED_RANGES.items():
        if not all(QtGui.QFontDatabase.supportsCharacter(font, code) for code in codes):
            return False, label
    return True, None


def _filter_supported_families(
    families: Iterable[str], source: str
) -> set[str]:
    """Filter *families* by glyph coverage and log skipped entries."""

    valid: set[str] = set()
    for family in families:
        ok, missing = _font_has_required_glyphs(family)
        if ok:
            valid.add(family)
        else:
            logger.warning(
                "Шрифт '%s' из '%s' пропущен: отсутствует %s набор символов",
                family,
                source,
                missing,
            )
    return valid


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
        message = (
            "Не удалось загрузить tkinter/customtkinter — шрифт Cattedrale не будет применён"
        )
        logger.warning(message)
        _show_error_dialog(message)
        return "Exo 2"

    if os.name != "nt" and not os.environ.get("DISPLAY"):
        message = (
            "Переменная окружения DISPLAY не установлена — шрифт Cattedrale не будет применён"
        )
        logger.warning(message)
        _show_error_dialog(message)
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
        """Apply the Exo 2 fallback and store it in global CONFIG."""

        fallback_family = "Exo 2"
        QGuiApplication.setFont(QFont(fallback_family))
        try:  # deferred import to avoid circular dependency
            from . import main as _main

            _main.CONFIG["font_family"] = fallback_family
            _main.CONFIG["header_font"] = fallback_family
            _main.CONFIG["sidebar_font"] = fallback_family
        except Exception:  # pragma: no cover - extremely defensive
            pass

    if not os.path.isdir(FONTS_DIR):
        _set_fallback()
        return

    extensions = {".ttf", ".otf", ".fon", ".ttc"}
    families: set[str] = set()

    font_path = os.path.join(FONTS_DIR, "Cattedrale[RUSbypenka220]-Regular.ttf")
    if tk and ctk and tkfont and (os.name == "nt" or os.environ.get("DISPLAY")):
        executor = ThreadPoolExecutor(max_workers=1)
        future = executor.submit(register_cattedrale, font_path)
        try:
            fam = future.result(timeout=2)
        except Exception:
            logger.exception(
                "Cattedrale registration failed or timed out; skipping"
            )
            fam = "Exo 2"
        finally:
            executor.shutdown(wait=False)
    else:
        logger.warning(
            "Tkinter unavailable or running headless; skipping Cattedrale registration"
        )
        fam = "Exo 2"
    if fam == "Exo 2":
        logger.warning("Cattedrale font not found, using fallback")
    preferred_cattedrale: str | None = None
    if fam != "Exo 2":
        fid = QtGui.QFontDatabase.addApplicationFont(font_path)
        if fid != -1:
            fams = QtGui.QFontDatabase.applicationFontFamilies(fid)
            valid = _filter_supported_families(fams, font_path)
            if valid:
                preferred_cattedrale = next(iter(valid))
                families.update(valid)
    if preferred_cattedrale is None and fam != "Exo 2":
        valid = _filter_supported_families({fam}, font_path)
        if valid:
            preferred_cattedrale = next(iter(valid))
            families.update(valid)
        else:
            fam = "Exo 2"

    whitelist_dirs = {os.path.normpath(os.path.join(FONTS_DIR, "Exo2"))}
    whitelist_files = {os.path.normpath(font_path)}

    def _iter_font_files() -> Iterable[str]:
        for entry in sorted(whitelist_files):
            normalized = os.path.normpath(entry)
            if normalized == os.path.normpath(font_path):
                continue
            if os.path.isfile(entry):
                yield entry
        for directory in sorted(whitelist_dirs):
            if not os.path.isdir(directory):
                continue
            for root, _dirs, files in os.walk(directory):
                for name in sorted(files):
                    yield os.path.join(root, name)

    for path in _iter_font_files():
        ext = os.path.splitext(path)[1].lower()
        if ext not in extensions:
            continue
        old_fams = set(QtGui.QFontDatabase.families())
        new_fams: set[str] = set()
        fid = (
            QtGui.QFontDatabase.addApplicationFont(path)
            if ext != ".fon"
            else -1
        )

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

        valid = _filter_supported_families(new_fams, path)
        families.update(valid)

    if "Exo 2" not in families:
        logger.error("Font 'Exo 2' not registered")
        _set_fallback()
        return

    # Success – ensure CONFIG reflects the available family and update header/sidebar
    try:  # pragma: no cover - defensive
        from . import main as _main
        from . import theme_manager

        _main.CONFIG.setdefault("font_family", "Exo 2")
        target_family = preferred_cattedrale if preferred_cattedrale else fam
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
