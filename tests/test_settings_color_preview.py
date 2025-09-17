import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets

import app.main as main


def test_color_preview_border_uses_updated_accent(tmp_path):
    main.CONFIG_PATH = str(tmp_path / "config.json")
    main.CONFIG = main.load_config()
    main.config.CONFIG = main.CONFIG

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    dlg = main.SettingsDialog()

    try:
        red_index = next(
            i
            for i, (_, color) in enumerate(dlg._preset_colors)
            if color.name().lower() == "#ff5555"
        )
        dlg._on_accent_changed(red_index)
        app.processEvents()

        style = dlg.btn_workspace.styleSheet()
        assert "border:2px solid #ff5555" in style
    finally:
        dlg.close()
        app.quit()
