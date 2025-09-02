import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets, QtGui

import resources
resources.register_fonts = lambda: None

import app.main as main

def test_lbl_month_font_and_border_on_theme_change():
    main.CONFIG["header_font"] = "Arial"
    main.CONFIG["theme"] = "light"
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    window = main.MainWindow()
    window.topbar.lbl_month.setFont(QtGui.QFont("Times"))
    window.topbar.lbl_month.setStyleSheet("border:1px solid red;")
    window.apply_theme()
    window.topbar.update_labels()
    assert window.topbar.lbl_month.font().family() == "Arial"
    assert "border:none" in window.topbar.lbl_month.styleSheet().replace(" ", "")
    window.close()
    app.quit()
