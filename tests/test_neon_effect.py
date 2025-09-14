import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets

import resources
resources.register_fonts = lambda: None

import app.main as main
from app.effects import apply_neon_effect

def test_lbl_month_border_reset_after_neon_effect():
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    window = main.MainWindow()
    label = window.topbar.lbl_month
    apply_neon_effect(label, True, config=main.CONFIG)
    apply_neon_effect(label, False, config=main.CONFIG)
    assert "border:none" in label.styleSheet().replace(" ", "")
    window.close()
    app.quit()
