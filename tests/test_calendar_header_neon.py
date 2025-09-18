import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets

import app.main as main
from app.effects import FixedDropShadowEffect


def test_calendar_header_neon_updates_with_accent():
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    table = main.ExcelCalendarTable()
    header = table.horizontalHeader()
    original_accent = main.CONFIG.get("accent_color")
    try:
        new_accent = "#fa5bff"
        main.CONFIG["accent_color"] = new_accent
        table.apply_theme()
        style = header.styleSheet().lower()
        assert new_accent.lower() in style
        effect = header.graphicsEffect()
        assert isinstance(effect, FixedDropShadowEffect)
    finally:
        main.CONFIG["accent_color"] = original_accent or "#39ff14"
        table.deleteLater()
        app.quit()
