import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets

import resources
resources.register_fonts = lambda: None

import app.main as main


def test_lbl_month_keeps_neon_during_and_after_settings(monkeypatch):
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    main.CONFIG["neon"] = True
    monkeypatch.setattr(main, "load_config", lambda: main.CONFIG)
    window = main.MainWindow()

    def fake_exec(self):
        label = window.topbar.lbl_month
        assert getattr(label, "_neon_effect", None) is not None
        self.settings_changed.emit()
        return 0

    monkeypatch.setattr(main.SettingsDialog, "exec", fake_exec)
    window.open_settings_dialog()
    label = window.topbar.lbl_month
    assert getattr(label, "_neon_effect", None) is not None
    window.close()
    app.quit()
