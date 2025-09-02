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


def test_sidebar_font_persists(tmp_path):
    main.CONFIG["sidebar_font"] = "Arial"
    main.CONFIG_PATH = str(tmp_path / "config.json")
    with open(main.CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(main.CONFIG, f)

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    window = main.MainWindow()
    window.apply_settings()

    first_btn = window.sidebar.buttons[0]
    assert first_btn.font().family() == "Arial"

    dlg = main.SettingsDialog(window)
    dlg.font_sidebar.setCurrentFont(QtGui.QFont("DejaVu Serif"))
    assert first_btn.font().family() == "DejaVu Serif"

    window.apply_settings()
    assert first_btn.font().family() == "DejaVu Serif"

    dlg.close()
    window.close()
    app.quit()
