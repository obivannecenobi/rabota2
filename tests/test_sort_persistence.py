import json
import os
import sys
from pathlib import Path

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6 import QtCore, QtWidgets
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

import resources  # noqa: E402

resources.register_fonts = lambda: None

import app.main as main  # noqa: E402


def _prepare_environment(tmp_path, monkeypatch):
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path))
    base = tmp_path / "data"
    main.BASE_SAVE_PATH = str(base)
    main.CONFIG["save_path"] = main.BASE_SAVE_PATH
    return QtWidgets.QApplication.instance() or QtWidgets.QApplication([])


def test_stats_dialog_restores_sort(tmp_path, monkeypatch):
    app = _prepare_environment(tmp_path, monkeypatch)

    stats_path = Path(main.stats_dir(2024)) / "2024.json"
    stats_path.parent.mkdir(parents=True, exist_ok=True)
    with open(stats_path, "w", encoding="utf-8") as fh:
        json.dump(
            {
                "1": [
                    {
                        "work": "Alpha",
                        "status": "",
                        "adult": False,
                        "total_chapters": 5,
                        "chars_per_chapter": 1000,
                        "planned": 4,
                        "chapters": 2,
                        "progress": 50.0,
                        "release": "",
                        "profit": 10.0,
                        "ads": 1.0,
                        "views": 100,
                        "likes": 5,
                        "thanks": 2,
                        "chars": 2000,
                    },
                    {
                        "work": "Beta",
                        "status": "",
                        "adult": False,
                        "total_chapters": 8,
                        "chars_per_chapter": 800,
                        "planned": 6,
                        "chapters": 5,
                        "progress": 80.0,
                        "release": "",
                        "profit": 12.0,
                        "ads": 1.5,
                        "views": 150,
                        "likes": 7,
                        "thanks": 3,
                        "chars": 4000,
                    },
                ]
            },
            fh,
            ensure_ascii=False,
            indent=2,
        )

    dlg = main.StatsDialog(2024, 1, None)
    app.processEvents()
    dlg.table_stats.sortByColumn(0, QtCore.Qt.SortOrder.DescendingOrder)
    app.processEvents()
    dlg.close()
    app.processEvents()

    dlg2 = main.StatsDialog(2024, 1, None)
    app.processEvents()
    header = dlg2.table_stats.horizontalHeader()
    assert header.sortIndicatorSection() == 0
    assert header.sortIndicatorOrder() == QtCore.Qt.SortOrder.DescendingOrder
    works = [
        dlg2.table_stats.item(row, 0).text()
        for row in range(dlg2.table_stats.rowCount())
        if dlg2.table_stats.item(row, 0) is not None
        and dlg2.table_stats.item(row, 0).text()
    ]
    assert works[:2] == ["Beta", "Alpha"]
    dlg2.close()
    app.quit()


def test_top_dialog_restores_sort(tmp_path, monkeypatch):
    app = _prepare_environment(tmp_path, monkeypatch)

    stats_path = Path(main.stats_dir(2024)) / "2024.json"
    stats_path.parent.mkdir(parents=True, exist_ok=True)
    with open(stats_path, "w", encoding="utf-8") as fh:
        json.dump(
            {
                "1": [
                    {
                        "work": "Alpha",
                        "status": "В процессе",
                        "total_chapters": 10,
                        "planned": 6,
                        "chapters": 4,
                        "progress": 40.0,
                        "release": "",
                        "chars": 3000,
                        "views": 100,
                        "profit": 100.0,
                        "ads": 15.0,
                        "likes": 10,
                        "thanks": 1,
                    },
                    {
                        "work": "Beta",
                        "status": "В процессе",
                        "total_chapters": 12,
                        "planned": 8,
                        "chapters": 7,
                        "progress": 70.0,
                        "release": "",
                        "chars": 4200,
                        "views": 160,
                        "profit": 150.0,
                        "ads": 10.0,
                        "likes": 15,
                        "thanks": 4,
                    },
                ]
            },
            fh,
            ensure_ascii=False,
            indent=2,
        )

    dlg = main.TopDialog(2024, None)
    app.processEvents()
    dlg.table.sortByColumn(0, QtCore.Qt.SortOrder.DescendingOrder)
    app.processEvents()
    dlg.close()
    app.processEvents()

    dlg2 = main.TopDialog(2024, None)
    app.processEvents()
    header = dlg2.table.horizontalHeader()
    assert header.sortIndicatorSection() == 0
    assert header.sortIndicatorOrder() == QtCore.Qt.SortOrder.DescendingOrder
    works = [
        dlg2.table.item(row, 1).text()
        for row in range(dlg2.table.rowCount())
        if dlg2.table.item(row, 1) is not None
        and dlg2.table.item(row, 1).text() not in {"", "Итого"}
    ]
    assert works[:2] == ["Beta", "Alpha"]
    dlg2.close()
    app.quit()

