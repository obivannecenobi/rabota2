import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets, QtGui

import resources
resources.register_fonts = lambda: None

import app.main as main


def test_calendar_table_theme_rebuilds_neon_without_css_accumulation():
    original_workspace = main.CONFIG.get("workspace_color")
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    window: main.MainWindow | None = None
    try:
        main.CONFIG["workspace_color"] = "#101010"
        window = main.MainWindow()
        window.apply_theme()

        table = window.table
        style_once = table.styleSheet()
        condensed = style_once.replace(" ", "").replace("\n", "")

        accent = QtGui.QColor(main.CONFIG.get("accent_color", "#39ff14"))
        subtle = QtGui.QColor(accent)
        subtle.setAlpha(90)
        subtle_r, subtle_g, subtle_b, subtle_a = subtle.getRgb()

        expected_subtle = f"border:1pxsolidrgba({subtle_r},{subtle_g},{subtle_b},{subtle_a})"
        assert expected_subtle in condensed
        assert "border-radius:16px" in condensed

        accent_name = accent.name()
        neon_border = f"border:1pxsolid{accent_name}"
        assert condensed.count(neon_border) == 1

        window.apply_theme()
        style_twice = window.table.styleSheet()
        assert style_twice == style_once
    finally:
        if original_workspace is not None:
            main.CONFIG["workspace_color"] = original_workspace
        if window is not None:
            window.close()
        if app is not None:
            app.quit()
