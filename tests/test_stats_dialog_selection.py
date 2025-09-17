import json
import os
import sys
from pathlib import Path

import pytest
from PySide6 import QtWidgets, QtCore, QtGui

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

import resources  # noqa: E402

resources.register_fonts = lambda: None

import app.main as main  # noqa: E402  pylint: disable=wrong-import-position


def _make_record(work: str, profit: float) -> dict[str, object]:
    return {
        "work": work,
        "status": "ongoing",
        "adult": False,
        "total_chapters": 10,
        "chars_per_chapter": 1000,
        "planned": 5,
        "chapters": 2,
        "progress": 20.0,
        "release": "2024-05",
        "profit": profit,
        "ads": 0.0,
        "views": 100,
        "likes": 3,
        "thanks": 1,
    }


def test_stats_dialog_selects_actual_record(tmp_path, monkeypatch):
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path))
    main.BASE_SAVE_PATH = str(tmp_path / "data")
    main.CONFIG["save_path"] = main.BASE_SAVE_PATH

    year, month = 2024, 5
    stats_path = Path(main.stats_dir(year)) / f"{year}.json"
    stats_path.parent.mkdir(parents=True, exist_ok=True)
    data = {str(month): [_make_record("Alpha", 100.0), _make_record("Beta", 200.0)]}
    stats_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    dialog = main.StatsDialog(year, month)
    dialog.show()
    app.processEvents()

    dialog.table_stats.sortItems(0, QtCore.Qt.DescendingOrder)
    app.processEvents()

    dialog.table_stats.selectRow(0)
    app.processEvents()

    assert dialog.current_index == 1
    assert dialog.form_stats.widgets["work"].text() == "Beta"
    assert dialog.form_stats.widgets["profit"].value() == pytest.approx(200.0)

    dialog.close()
    app.processEvents()
    app.quit()


def test_stats_dialog_table_style_uses_accent(tmp_path, monkeypatch):
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path))
    main.BASE_SAVE_PATH = str(tmp_path / "data")
    main.CONFIG["save_path"] = main.BASE_SAVE_PATH

    year, month = 2024, 6
    stats_path = Path(main.stats_dir(year)) / f"{year}.json"
    stats_path.parent.mkdir(parents=True, exist_ok=True)
    data = {str(month): [_make_record("Gamma", 50.0)]}
    stats_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    dialog = main.StatsDialog(year, month)
    dialog.show()
    app.processEvents()

    accent = QtGui.QColor(main.CONFIG.get("accent_color", "#39ff14")).name().lower()
    table_style = dialog.table_stats.styleSheet().lower()
    header_style = dialog.table_stats.horizontalHeader().styleSheet().lower()

    assert accent in table_style
    assert accent in header_style
    assert getattr(dialog.table_stats, "_neon_filter", None) is not None
    assert getattr(dialog.table_stats.horizontalHeader(), "_neon_filter", None) is not None

    dialog.close()
    app.processEvents()
    app.quit()
