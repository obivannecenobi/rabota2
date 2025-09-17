import os
import sys
import json
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets, QtGui

import resources

resources.register_fonts = lambda: None

import app.main as main


def test_accent_color_updates_spinbox_border(tmp_path):
    main.CONFIG_PATH = str(tmp_path / "config.json")
    with open(main.CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(main.CONFIG, f)

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    window = main.MainWindow()
    window.apply_settings()

    new_accent = "#ff8800"
    main.CONFIG["accent_color"] = new_accent
    with open(main.CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(main.CONFIG, f)

    window.apply_theme()

    style = window.topbar.spin_year.styleSheet().replace(" ", "").lower()
    expected = QtGui.QColor(new_accent).name().lower()
    assert f"border:1pxsolid{expected}" in style

    window.close()
    app.quit()
