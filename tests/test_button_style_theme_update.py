import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets

import resources
resources.register_fonts = lambda: None

import app.main as main
from widgets import StyledPushButton


def test_apply_settings_preserves_idle_neon(monkeypatch):
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    window = main.MainWindow()

    idle_button = StyledPushButton(window, **main.button_config())
    idle_button.setText("Idle")
    idle_button.setProperty("neon_selected", False)
    idle_button.show()

    monkeypatch.setitem(main.CONFIG, "neon", True)

    window.apply_settings()
    window.apply_settings()
    QtWidgets.QApplication.processEvents()

    effect = idle_button.graphicsEffect()
    assert effect is not None

    style = idle_button.styleSheet().lower()
    accent = main.CONFIG.get("accent_color", "#39ff14").lower()
    assert f"border:" in style
    assert accent in style

    idle_button.close()
    window.close()
    app.quit()
