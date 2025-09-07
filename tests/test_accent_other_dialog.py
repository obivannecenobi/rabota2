import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets, QtGui

import resources

resources.register_fonts = lambda: None

import app.main as main


def test_other_selection_always_opens_dialog(monkeypatch):
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    dlg = main.SettingsDialog()

    calls = []

    def fake_get_color(initial, *args, **kwargs):
        calls.append(True)
        # return some valid color
        return QtGui.QColor("#010203")

    monkeypatch.setattr(QtWidgets.QColorDialog, "getColor", fake_get_color)

    other_index = dlg.combo_accent.count() - 1

    # simulate selecting "Other" twice without changing index
    dlg.combo_accent.setCurrentIndex(other_index)
    dlg.combo_accent.activated.emit(other_index)
    dlg.combo_accent.activated.emit(other_index)

    assert len(calls) == 2

    dlg.close()
    app.quit()

