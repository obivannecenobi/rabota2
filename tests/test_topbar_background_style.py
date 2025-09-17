import os
import re
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets, QtGui

import resources

resources.register_fonts = lambda: None

import app.main as main


def test_topbar_background_stylesheet_overwrites_previous_colors():
    original_color = main.CONFIG.get("workspace_color")
    original_monochrome = main.CONFIG.get("monochrome")

    try:
        main.CONFIG["monochrome"] = False
        main.CONFIG["workspace_color"] = "#101010"

        app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        window = main.MainWindow()
        window.apply_theme()

        previous_colors: list[str] = []
        for color in ("#112233", "#445566"):
            main.CONFIG["workspace_color"] = color
            window.apply_theme()

            current_name = QtGui.QColor(color).name()
            stylesheet = window.topbar.styleSheet()

            assert stylesheet.count(current_name) == 1
            assert re.search(r"border-radius\s*:\s*16px", stylesheet) is not None
            for old in previous_colors:
                assert QtGui.QColor(old).name() not in stylesheet

            previous_colors.append(color)

        window.close()
        app.quit()
    finally:
        if original_color is not None:
            main.CONFIG["workspace_color"] = original_color
        else:
            main.CONFIG.pop("workspace_color", None)
        if original_monochrome is not None:
            main.CONFIG["monochrome"] = original_monochrome
        else:
            main.CONFIG.pop("monochrome", None)
