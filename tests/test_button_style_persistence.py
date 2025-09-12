import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets

import resources
resources.register_fonts = lambda: None

import app.main as main
from widgets import StyledToolButton


def test_gradient_and_neon_persist_across_buttons_and_dialog(monkeypatch):
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    window = main.MainWindow()
    window.sidebar.activate_button(window.sidebar.buttons[0])

    dlg = QtWidgets.QDialog()
    btn_dialog = StyledToolButton(dlg)
    btn_dialog.setText("Dlg")
    QtWidgets.QVBoxLayout(dlg).addWidget(btn_dialog)

    main.CONFIG["gradient_colors"] = ["#123456", "#654321"]
    main.CONFIG["neon"] = True
    monkeypatch.setattr(main, "load_config", lambda: main.CONFIG)

    window.apply_settings()

    grad = main.CONFIG["gradient_colors"]
    targets = [window.sidebar.buttons[0], window.topbar.btn_prev, btn_dialog]
    for b in targets:
        style = b.styleSheet().lower()
        assert f"stop:0 {grad[0].lower()}" in style
        assert f"stop:1 {grad[1].lower()}" in style
        assert b.graphicsEffect() is not None

    dlg.close()

    sidebar_btn = window.sidebar.buttons[0]
    style = sidebar_btn.styleSheet().lower()
    assert f"stop:0 {grad[0].lower()}" in style
    assert sidebar_btn.graphicsEffect() is not None

    window.close()
    app.quit()
