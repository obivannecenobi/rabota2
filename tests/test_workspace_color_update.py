import os
import sys
import json
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets, QtGui

import resources
resources.register_fonts = lambda: None

if not hasattr(QtGui.QFontDatabase, "supportsCharacter"):
    QtGui.QFontDatabase.supportsCharacter = staticmethod(lambda *args, **kwargs: True)

import app.main as main


def test_workspace_color_updates_bars_and_persists(tmp_path, monkeypatch):
    main.CONFIG["workspace_color"] = "#111111"
    main.CONFIG_PATH = str(tmp_path / "config.json")
    with open(main.CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(main.CONFIG, f)

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    window = main.MainWindow()
    window.apply_settings()
    dlg = main.SettingsDialog(window)

    new_color = QtGui.QColor("#222222")
    monkeypatch.setattr(QtWidgets.QColorDialog, "getColor", lambda *a, **k: new_color)

    dlg.choose_workspace_color()

    assert window.topbar.palette().color(QtGui.QPalette.Window).name() == new_color.name()
    assert f"background-color:{new_color.name()}" in window.statusBar().styleSheet()

    dlg.close()
    assert window.topbar.palette().color(QtGui.QPalette.Window).name() == new_color.name()
    assert f"background-color:{new_color.name()}" in window.statusBar().styleSheet()

    window.close()
    app.quit()


def test_topbar_month_label_tracks_accent_color(monkeypatch):
    _ = monkeypatch
    original_accent = main.CONFIG.get("accent_color")
    original_monochrome = main.CONFIG.get("monochrome", False)
    original_saturation = main.CONFIG.get("mono_saturation", 100)

    try:
        main.CONFIG["monochrome"] = False
        main.CONFIG["mono_saturation"] = original_saturation
        first_color = "#123456"
        second_color = "#abcdef"
        main.CONFIG["accent_color"] = first_color

        app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        topbar = main.TopBar()
        topbar.apply_style()

        main.CONFIG["accent_color"] = second_color
        topbar.apply_style()

        expected = QtGui.QColor(second_color).name().lower()
        palette_color = (
            topbar.lbl_month.palette()
            .color(QtGui.QPalette.WindowText)
            .name()
            .lower()
        )

        assert palette_color == expected
        assert expected in topbar.lbl_month.styleSheet().lower()
    finally:
        main.CONFIG["accent_color"] = original_accent
        main.CONFIG["monochrome"] = original_monochrome
        main.CONFIG["mono_saturation"] = original_saturation
        if "topbar" in locals():
            topbar.deleteLater()
        if "app" in locals():
            app.quit()
