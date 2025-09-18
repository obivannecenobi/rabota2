import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtCore, QtWidgets, QtGui

import resources
resources.register_fonts = lambda: None

import app.main as main
from widgets import StyledToolButton, StyledPushButton


def test_gradient_and_neon_persist_across_buttons_and_dialog(monkeypatch):
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    window = main.MainWindow()
    window.sidebar.activate_button(window.sidebar.buttons[0])

    dlg = QtWidgets.QDialog()
    btn_dialog = StyledToolButton(dlg, **main.button_config())
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

    # все кнопки сохраняют неоновый контур (selected усиливает эффект)
    assert window.sidebar.buttons[0].graphicsEffect() is not None
    assert window.topbar.btn_prev.graphicsEffect() is not None
    assert btn_dialog.graphicsEffect() is not None

    dlg.close()

    # Reopen a dialog to ensure sidebar retains updated gradient
    dlg = QtWidgets.QDialog()
    btn_dialog = StyledToolButton(dlg, **main.button_config())
    btn_dialog.setText("Dlg")
    QtWidgets.QVBoxLayout(dlg).addWidget(btn_dialog)
    window.apply_settings()

    sidebar_btn = window.sidebar.buttons[0]
    style = sidebar_btn.styleSheet().lower()
    assert f"stop:0 {grad[0].lower()}" in style
    assert sidebar_btn.graphicsEffect() is not None

    dlg.close()
    window.close()
    app.quit()


def test_sidebar_button_spacing_and_style_toggle(monkeypatch):
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    monkeypatch.setitem(main.CONFIG, "sidebar_collapsed", False)

    window = main.MainWindow()
    sidebar = window.sidebar

    labels = [btn.text() for btn in sidebar.buttons]
    assert all(label == label.strip() for label in labels)

    sidebar.set_collapsed(True)
    for btn in sidebar.buttons:
        assert btn.toolButtonStyle() == QtCore.Qt.ToolButtonIconOnly

    sidebar.set_collapsed(False)
    for original, btn in zip(labels, sidebar.buttons):
        assert btn.toolButtonStyle() == QtCore.Qt.ToolButtonTextBesideIcon
        assert btn.text() == original
        assert not btn.text().startswith(" ")

    window.close()
    app.quit()


def test_sidebar_frame_style_uses_accent_border(monkeypatch):
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    monkeypatch.setitem(main.CONFIG, "sidebar_collapsed", False)
    monkeypatch.setitem(main.CONFIG, "accent_color", "#ff44aa")
    monkeypatch.setitem(main.CONFIG, "sidebar_color", "#101010")

    window = main.MainWindow()
    window.apply_palette()
    sidebar = window.sidebar

    def current_style() -> str:
        return sidebar.styleSheet().lower()

    initial_style = current_style()
    assert "border-radius" in initial_style
    assert "#ff44aa" in initial_style

    monkeypatch.setitem(main.CONFIG, "accent_color", "#33cc55")
    window.apply_palette()
    updated_style = current_style()
    assert "#33cc55" in updated_style
    assert "#ff44aa" not in updated_style

    sidebar.set_collapsed(True)
    QtWidgets.QApplication.processEvents()
    collapsed_style = current_style()
    assert "border-radius" in collapsed_style
    assert "#33cc55" in collapsed_style

    sidebar.set_collapsed(False)
    QtWidgets.QApplication.processEvents()
    expanded_style = current_style()
    assert "border-radius" in expanded_style
    assert "#33cc55" in expanded_style

    window.close()
    app.quit()


def test_push_button_retains_neon_after_leave_event():
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    button = StyledPushButton()
    button.setText("Idle")
    button.resize(120, 40)
    button.show()
    QtWidgets.QApplication.processEvents()

    enter_event = QtGui.QEnterEvent(
        QtCore.QPointF(),
        QtCore.QPointF(),
        QtCore.QPointF(),
    )
    leave_event = QtCore.QEvent(QtCore.QEvent.Leave)
    QtWidgets.QApplication.sendEvent(button, enter_event)
    QtWidgets.QApplication.sendEvent(button, leave_event)

    assert button.graphicsEffect() is not None

    button.close()
    app.quit()
