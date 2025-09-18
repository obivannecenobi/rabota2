import json
import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets, QtGui

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


def test_resolve_font_config_rejects_unsupported(monkeypatch, tmp_path):
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    monkeypatch.setattr(main, "CONFIG_PATH", str(tmp_path / "config.json"))
    main.CONFIG = main.load_config()
    main.config.CONFIG = main.CONFIG
    main.CONFIG.update(
        {
            "header_font": "Unsupported",
            "text_font": "Unsupported",
            "sidebar_font": "Unsupported",
        }
    )

    monkeypatch.setattr(main, "ensure_font_registered", lambda family, parent=None: family)
    monkeypatch.setattr(
        resources,
        "ensure_supported_family",
        lambda family, source, fallback="Exo 2": (
            ("Exo 2", "кириллический") if family == "Unsupported" else (family, None)
        ),
    )

    header, text = main.resolve_font_config()

    try:
        assert header == "Exo 2"
        assert text == "Exo 2"
        assert main.CONFIG["header_font"] == "Exo 2"
        assert main.CONFIG["text_font"] == "Exo 2"
        assert main.CONFIG["sidebar_font"] == "Exo 2"
        assert main.CONFIG["font_family"] == "Exo 2"
        with open(main.CONFIG_PATH, "r", encoding="utf-8") as fh:
            stored = json.load(fh)
        assert stored["header_font"] == "Exo 2"
        assert stored["text_font"] == "Exo 2"
        assert stored["sidebar_font"] == "Exo 2"
        assert stored["font_family"] == "Exo 2"
    finally:
        app.quit()


def test_settings_dialog_restores_exo_on_invalid_selection(monkeypatch, tmp_path):
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    monkeypatch.setattr(main, "CONFIG_PATH", str(tmp_path / "config.json"))
    main.CONFIG = main.load_config()
    main.config.CONFIG = main.CONFIG

    monkeypatch.setattr(
        resources,
        "filter_supported_families",
        lambda families, source, emit_warnings=True: {"Exo 2"},
    )
    monkeypatch.setattr(
        resources,
        "ensure_supported_family",
        lambda family, source, fallback="Exo 2": (
            ("Exo 2", None) if family != "Exo 2" else (family, None)
        ),
    )

    def fake_details(family):
        if family == "FakeFont":
            return False, "кириллический"
        return True, None

    monkeypatch.setattr(resources, "family_support_details", fake_details)

    warnings: list[str] = []
    monkeypatch.setattr(
        QtWidgets.QMessageBox,
        "warning",
        lambda parent, title, text: warnings.append(text),
    )

    dialog = main.SettingsDialog()

    try:
        dialog._handle_font_combo_changed(
            "text_font", dialog.font_text, QtGui.QFont("FakeFont")
        )
        assert dialog.font_text.currentFont().family() == "Exo 2"
        assert main.CONFIG["text_font"] == "Exo 2"
        assert main.CONFIG["font_family"] == "Exo 2"
        assert warnings and "FakeFont" in warnings[0]
    finally:
        dialog.deleteLater()
        app.quit()
