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


def test_day_label_font_updates_immediately(tmp_path):
    main.CONFIG["header_font"] = "Arial"
    main.CONFIG_PATH = str(tmp_path / "config.json")
    with open(main.CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(main.CONFIG, f)

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    window = main.MainWindow()
    dlg = main.SettingsDialog(window)

    dlg.font_header.setCurrentFont(QtGui.QFont("DejaVu Serif"))
    lbl = next(iter(window.table.day_labels.values()))

    assert lbl.font().family() == "DejaVu Serif"

    dlg.close()
    window.close()
    app.quit()


def test_header_font_persists_after_month_change(tmp_path):
    main.CONFIG["header_font"] = "Arial"
    main.CONFIG_PATH = str(tmp_path / "config.json")
    with open(main.CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(main.CONFIG, f)

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    window = main.MainWindow()
    dlg = main.SettingsDialog(window)
    dlg.font_header.setCurrentFont(QtGui.QFont("DejaVu Serif"))
    dlg.close()

    window.next_month()
    lbl = next(iter(window.table.day_labels.values()))

    assert lbl.font().family() == "DejaVu Serif"
    assert window.topbar.lbl_month.font().family() == "DejaVu Serif"

    window.close()
    app.quit()

