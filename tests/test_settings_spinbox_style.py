import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets

import app.main as main


def _prepare_config(tmp_path):
    main.CONFIG_PATH = str(tmp_path / "config.json")
    main.CONFIG = main.load_config()
    main.config.CONFIG = main.CONFIG


def test_spin_day_rows_updates_accent_color(tmp_path):
    _prepare_config(tmp_path)

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    dlg = main.SettingsDialog()

    try:
        initial_style = dlg.spin_day_rows.styleSheet().lower()
        assert main.CONFIG["accent_color"].lower() in initial_style

        # switch to the preset "Красный" (index 1) and ensure the style updates
        dlg._on_accent_changed(1)
        updated_style = dlg.spin_day_rows.styleSheet().lower()
        assert "#ff5555" in updated_style
    finally:
        dlg.close()
        app.quit()
