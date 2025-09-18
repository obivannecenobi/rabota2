import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets

import resources

resources.register_fonts = lambda: None

import app.main as main


def test_inner_table_style_updates_with_accent(monkeypatch):
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    monkeypatch.setattr(main.ExcelCalendarTable, "load_month_data", lambda self, y, m: None)

    table = main.ExcelCalendarTable()
    inner = table._create_inner_table()
    table.cell_tables[(0, 0)] = inner

    old_accent = str(main.CONFIG.get("accent_color", "#39ff14"))
    initial_style = inner.styleSheet().lower()
    assert old_accent.lower() in initial_style

    new_accent = "#ff55aa"
    main.CONFIG["accent_color"] = new_accent

    try:
        table.apply_theme()
        updated_style = inner.styleSheet().lower()
        assert new_accent in updated_style
        assert getattr(inner, "_neon_effect", None) is not None
    finally:
        main.CONFIG["accent_color"] = old_accent
        table.deleteLater()
        inner.deleteLater()
        app.quit()
