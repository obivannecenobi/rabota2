import os, sys, json
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets, QtGui

import resources
resources.register_fonts = lambda: None

import app.main as main


def test_accent_color_persists_across_restart(tmp_path, monkeypatch):
    # use temporary config file
    main.CONFIG_PATH = str(tmp_path / "config.json")
    with open(main.CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(main.CONFIG, f)

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    window = main.MainWindow()
    window.apply_settings()

    dlg = main.SettingsDialog(window)
    dlg.settings_changed.connect(window._on_settings_changed)

    new_color = QtGui.QColor("#abcdef")
    monkeypatch.setattr(QtWidgets.QColorDialog, "getColor", lambda *a, **k: new_color)

    other_index = dlg.combo_accent.count() - 1
    dlg.combo_accent.setCurrentIndex(other_index)
    dlg.combo_accent.activated.emit(other_index)

    # verify highlight and button style update
    highlight = app.palette().color(QtGui.QPalette.Highlight).name()
    assert highlight == new_color.name().lower()

    btn = window.sidebar.buttons[0]
    btn._apply_hover(True)
    assert new_color.name().lower() in btn.styleSheet().lower()
    btn._apply_hover(False)

    # close and restart application
    window.close()
    app.processEvents()

    main.CONFIG = main.load_config()
    window2 = main.MainWindow()
    window2.apply_settings()

    highlight2 = app.palette().color(QtGui.QPalette.Highlight).name()
    assert highlight2 == new_color.name().lower()

    window2.sidebar.activate_button(window2.sidebar.buttons[0])
    assert new_color.name().lower() in window2.sidebar.buttons[0].styleSheet().lower()

    window2.close()
    app.quit()
