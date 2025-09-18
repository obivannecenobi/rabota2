import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets

import resources

resources.register_fonts = lambda: None

import app.main as main


def test_sidebar_button_stays_selected_after_dialog(monkeypatch):
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    window = main.MainWindow()

    btn = window.sidebar.btn_inputs
    assert btn is not None
    window.sidebar.activate_button(btn)

    class DummyDialog:
        def __init__(self, *args, **kwargs):
            pass

        def exec(self):
            return 0

    monkeypatch.setattr(main, "StatsDialog", DummyDialog)

    window.open_input_dialog()

    assert window.sidebar.last_active_button is btn
    assert btn.property("neon_selected") is True
    for other in window.sidebar.buttons:
        if other is btn:
            continue
        assert other.property("neon_selected") is False

    window.sidebar.set_collapsed(True)
    window.sidebar.set_collapsed(False)
    assert window.sidebar.last_active_button is btn

    window.close()
    app.quit()
