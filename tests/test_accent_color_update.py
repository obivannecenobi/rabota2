import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
# ensure app package importable
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets, QtGui

import resources
resources.register_fonts = lambda: None

import app.main as main


def test_sidebar_updates_on_accent_change(monkeypatch):
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    window = main.MainWindow()
    window.apply_settings()

    calls = []

    def fake_apply_style(self, neon, accent, sidebar_color):
        # record the accent and palette highlight at call time
        highlight = QtWidgets.QApplication.instance().palette().color(
            QtGui.QPalette.Highlight
        ).name()
        calls.append((accent.name(), highlight))

    monkeypatch.setattr(window.sidebar, "apply_style", fake_apply_style)

    new_color = "#ff0000"
    main.CONFIG["accent_color"] = new_color
    window.apply_palette()

    assert calls, "apply_style was not called"
    accent_name, highlight_name = calls[-1]
    assert accent_name == new_color.lower()
    assert highlight_name == new_color.lower()

    window.close()
    app.quit()
