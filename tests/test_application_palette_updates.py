import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets, QtGui

import resources

resources.register_fonts = lambda: None

import app.main as main


def test_apply_theme_updates_global_palette():
    previous_color = main.CONFIG.get("workspace_color", None)
    _monochrome_sentinel = object()
    previous_monochrome = main.CONFIG.get("monochrome", _monochrome_sentinel)

    app = QtWidgets.QApplication.instance()
    created_app = False
    if app is None:
        app = QtWidgets.QApplication([])
        created_app = True

    window = main.MainWindow()
    window.apply_settings()

    try:
        new_color = "#345678"
        main.CONFIG["monochrome"] = False
        main.CONFIG["workspace_color"] = new_color

        window.apply_theme()

        applied_color = QtWidgets.QApplication.palette().color(QtGui.QPalette.Window).name()
        assert applied_color == QtGui.QColor(new_color).name()
    finally:
        window.close()
        if previous_color is not None:
            main.CONFIG["workspace_color"] = previous_color
        else:
            main.CONFIG.pop("workspace_color", None)
        if previous_monochrome is _monochrome_sentinel:
            main.CONFIG.pop("monochrome", None)
        else:
            main.CONFIG["monochrome"] = previous_monochrome
        if created_app:
            app.quit()
