import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets, QtGui

import resources
resources.register_fonts = lambda: None

import app.main as main


def test_calendar_status_row_colors_follow_workspace():
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    main.CONFIG["workspace_color"] = "#111111"
    window = main.MainWindow()
    window.apply_theme()
    style_before = window.table.styleSheet().replace(" ", "")
    assert "#111111" in style_before

    # change color and reapply
    main.CONFIG["workspace_color"] = "#222222"
    window.apply_theme()
    style_after = window.table.styleSheet().replace(" ", "")
    assert "#222222" in style_after and "#111111" not in style_after
    assert "QTableWidget::item" in style_after
    text_color = window.palette().color(QtGui.QPalette.WindowText).name()
    assert f"color:{text_color}" in style_after

    inner_tbl = next(iter(window.table.cell_tables.values()))
    inner_style = inner_tbl.styleSheet().replace(" ", "")
    assert "#222222" in inner_style
    window.close()
    app.quit()
