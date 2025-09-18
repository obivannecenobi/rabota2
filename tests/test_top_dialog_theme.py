import json
import os
import sys
from pathlib import Path

from PySide6 import QtGui, QtWidgets
import shiboken6

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

import resources  # noqa: E402

resources.register_fonts = lambda: None

import app.main as main  # noqa: E402  pylint: disable=wrong-import-position


def _make_top_record(work: str, profit: float) -> dict[str, object]:
    return {
        "work": work,
        "status": "active",
        "total_chapters": 10,
        "planned": 5,
        "chapters": 3,
        "progress": 30.0,
        "release": "2024-01",
        "chars": 1200,
        "views": 150,
        "profit": profit,
        "ads": 15.0,
        "likes": 4,
        "thanks": 2,
    }


def test_top_dialog_table_uses_accent_color(tmp_path, monkeypatch):
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path))
    main.BASE_SAVE_PATH = str(tmp_path / "data")
    main.CONFIG["save_path"] = main.BASE_SAVE_PATH

    accent_color = "#123abc"
    workspace_color = "#202124"
    main.CONFIG["accent_color"] = accent_color
    main.CONFIG["workspace_color"] = workspace_color

    year = 2024
    stats_path = Path(main.stats_dir(year)) / f"{year}.json"
    stats_path.parent.mkdir(parents=True, exist_ok=True)
    stats_data = {"1": [_make_top_record("Project Alpha", 250.0)]}
    stats_path.write_text(json.dumps(stats_data, ensure_ascii=False, indent=2), encoding="utf-8")

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    dialog = main.TopDialog(year)
    dialog.show()
    app.processEvents()

    headers = [
        dialog.table.horizontalHeaderItem(i).text()
        for i in range(dialog.table.columnCount())
    ]
    assert "Запланировано" in headers
    assert "Запланированно" not in headers

    accent = QtGui.QColor(accent_color).name().lower()
    table_style = dialog.table.styleSheet().lower()
    header_style = dialog.table.horizontalHeader().styleSheet().lower()

    assert accent in table_style
    assert accent in header_style
    assert getattr(dialog.table, "_neon_filter", None) is not None
    assert getattr(dialog.table.horizontalHeader(), "_neon_filter", None) is not None

    dialog.close()
    app.processEvents()
    app.quit()


def test_top_dialog_refresh_updates_filter_controls(tmp_path, monkeypatch):
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path))
    main.BASE_SAVE_PATH = str(tmp_path / "data")
    main.CONFIG["save_path"] = main.BASE_SAVE_PATH

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    dialog = main.TopDialog(2024)
    dialog.show()
    app.processEvents()

    new_accent = "#f5429b"
    new_workspace = "#14161c"
    monkeypatch.setitem(main.CONFIG, "accent_color", new_accent)
    monkeypatch.setitem(main.CONFIG, "workspace_color", new_workspace)

    dialog.refresh_theme()
    app.processEvents()

    accent = QtGui.QColor(new_accent).name().lower()
    style = dialog.spin_year.styleSheet().lower()
    assert f"border:1px solid {accent}" in style

    effect = dialog.btn_calc.graphicsEffect()
    assert effect is not None and shiboken6.isValid(effect)

    dialog.close()
    app.processEvents()
    app.quit()
