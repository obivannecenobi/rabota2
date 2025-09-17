import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets

import app.main as main
import resources


def test_register_fonts_falls_back_to_exo2(monkeypatch, tmp_path):
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    main.CONFIG_PATH = str(tmp_path / "config.json")
    main.CONFIG = main.load_config()
    main.config.CONFIG = main.CONFIG

    monkeypatch.setattr(resources, "_filter_supported_families", lambda families, source: set())

    resources.register_fonts()

    try:
        assert main.CONFIG["font_family"] == "Exo 2"
        assert main.CONFIG["header_font"] == "Exo 2"
        assert main.CONFIG["sidebar_font"] == "Exo 2"
    finally:
        app.quit()
