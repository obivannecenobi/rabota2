import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets

import app.main as main


def test_settings_dialog_uses_exo2_by_default(tmp_path):
    main.CONFIG_PATH = str(tmp_path / "config.json")
    main.CONFIG = main.load_config()
    main.config.CONFIG = main.CONFIG

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    dlg = main.SettingsDialog()

    try:
        assert dlg.font_header.currentFont().family() == "Exo 2"
        assert dlg.font_text.currentFont().family() == "Exo 2"
        assert dlg.font_sidebar.currentFont().family() == "Exo 2"
    finally:
        dlg.close()
        app.quit()
