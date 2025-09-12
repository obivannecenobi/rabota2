import os
import sys
import json
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets

import resources
resources.register_fonts = lambda: None

import app.main as main


def test_year_spinbox_border_persists(tmp_path):
    # Use temporary config file
    main.CONFIG_PATH = str(tmp_path / "config.json")
    with open(main.CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(main.CONFIG, f)

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    window = main.MainWindow()
    window.apply_settings()

    # change workspace color and simulate settings save
    main.CONFIG["workspace_color"] = "#123456"
    with open(main.CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(main.CONFIG, f)

    window._on_settings_changed()

    style = window.topbar.spin_year.styleSheet().replace(" ", "")
    assert "border-radius:6px" in style
    assert "border:1pxsolidrgba(255,255,255,0.2)" in style

    window.close()
    app.quit()
