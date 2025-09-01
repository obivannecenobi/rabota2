"""Application utility helpers."""

from __future__ import annotations

import json
import logging
import os
from typing import Dict, Tuple

from PySide6 import QtWidgets, QtGui


LOGGER = logging.getLogger(__name__)

ASSETS_DIR = os.path.join(os.path.dirname(__file__), "..", "assets")
DATA_DIR = os.path.join(os.path.dirname(__file__), "..", "data")
CONFIG_PATH = os.path.join(DATA_DIR, "config.json")

DEFAULT_CONFIG: Dict[str, object] = {
    "neon": False,
    "neon_size": 10,
    "neon_thickness": 1,
    "neon_intensity": 255,
    "accent_color": "#39ff14",
    "gradient_colors": ["#39ff14", "#2d7cdb"],
    "gradient_angle": 0,
    "glass_effect": "",
    "glass_opacity": 0.5,
    "glass_blur": 10,
    "glass_enabled": False,
    "header_font": "Exo 2",
    "text_font": "Exo 2",
    "save_path": DATA_DIR,
    "day_rows": 6,
    "workspace_color": "#1e1e21",
    "sidebar_color": "#1f1f23",
    "monochrome": False,
    "mono_saturation": 0.5,
    "theme": "dark",
}


def _write_config(data: Dict[str, object]) -> None:
    os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)
    with open(CONFIG_PATH, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)


def load_config() -> Dict[str, object]:
    """Load JSON configuration, recreating defaults on corruption."""

    config = DEFAULT_CONFIG.copy()
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            if isinstance(data, dict):
                config.update({k: v for k, v in data.items() if v is not None})
            save_path = config.get("save_path")
            if isinstance(save_path, str) and not os.path.isabs(save_path):
                config["save_path"] = os.path.abspath(
                    os.path.join(os.path.dirname(CONFIG_PATH), save_path)
                )
        except json.JSONDecodeError:
            LOGGER.error("Config file is corrupted, recreating default")
            _write_config(config)
        except Exception:  # pragma: no cover - unexpected errors
            LOGGER.exception("Failed to load config; using defaults")
    else:
        _write_config(config)
    return config


CONFIG = load_config()
BASE_SAVE_PATH = os.path.abspath(CONFIG.get("save_path", DATA_DIR))


def ensure_font_registered(
    family: str, parent: QtWidgets.QWidget | None = None
) -> str:
    """Ensure *family* font is available and return its name."""

    if family in QtGui.QFontDatabase.families():
        return family

    file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
        parent,
        "Выберите файл шрифта",
        "",
        "Font Files (*.ttf *.otf)",
    )
    if file_path:
        fid = QtGui.QFontDatabase.addApplicationFont(file_path)
        if fid != -1:
            fams = QtGui.QFontDatabase.applicationFontFamilies(fid)
            if fams:
                return fams[0]
        LOGGER.error("Failed to load font from '%s'", file_path)

    LOGGER.warning("Falling back to system font 'Sans Serif'")
    return "Sans Serif"


def resolve_font_config(
    parent: QtWidgets.QWidget | None = None,
) -> Tuple[str, str]:
    """Resolve configured fonts and persist changes."""

    header = ensure_font_registered(CONFIG.get("header_font", "Exo 2"), parent)
    text = ensure_font_registered(CONFIG.get("text_font", "Exo 2"), parent)
    changed = False
    if header != CONFIG.get("header_font"):
        CONFIG["header_font"] = header
        changed = True
    if text != CONFIG.get("text_font"):
        CONFIG["text_font"] = text
        changed = True
    if changed:
        _write_config(CONFIG)
    return header, text

