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


class DummyTable:
    def __init__(self):
        self.applied = False

    def apply_theme(self):
        self.applied = True


class DummyTopBar(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.color = None

    def apply_background(self, color):
        self.color = QtGui.QColor(color)


class DummyParent(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.table = DummyTable()
        self.topbar = DummyTopBar()


def test_workspace_color_updates_immediately(tmp_path, monkeypatch):
    main.CONFIG["workspace_color"] = "#111111"
    main.CONFIG_PATH = str(tmp_path / "config.json")
    with open(main.CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(main.CONFIG, f)

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    parent = DummyParent()
    dlg = main.SettingsDialog(parent)

    new_color = QtGui.QColor("#222222")
    monkeypatch.setattr(QtWidgets.QColorDialog, "getColor", lambda *a, **k: new_color)

    dlg.choose_workspace_color()

    assert parent.table.applied
    assert parent.topbar.color.name() == new_color.name()

    dlg.close()
    app.quit()
