import os
import sys
from pathlib import Path

import pytest
from PySide6 import QtWidgets, QtCore

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

import resources  # noqa: E402

resources.register_fonts = lambda: None

import app.main as main  # noqa: E402


@pytest.mark.parametrize(
    "factory",
    [
        lambda parent: main.StatsDialog(2024, 1, parent),
        lambda parent: main.AnalyticsDialog(2024, parent),
        lambda parent: main.ReleaseDialog(2024, 1, [], parent),
        lambda parent: main.TopDialog(2024, parent),
        lambda parent: main.SettingsDialog(parent),
    ],
)
def test_dialog_geometry_persist(tmp_path, monkeypatch, factory):
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path))
    main.BASE_SAVE_PATH = str(tmp_path / "data")
    main.CONFIG["save_path"] = main.BASE_SAVE_PATH
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    dlg = factory(None)
    size = QtCore.QSize(640, 480)
    dlg.resize(size)
    dlg.show()
    app.processEvents()
    expected = dlg.size()
    dlg.close()
    app.processEvents()
    dlg2 = factory(None)
    dlg2.show()
    app.processEvents()
    assert dlg2.size() == expected
    dlg2.close()
    app.quit()


def test_calendar_columns_persist(tmp_path, monkeypatch):
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path))
    main.BASE_SAVE_PATH = str(tmp_path / "data")
    main.CONFIG["save_path"] = main.BASE_SAVE_PATH
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    window = main.MainWindow()
    # Adjust inner day table column widths via header resizing
    inner = next(iter(window.table.cell_tables.values()))
    header = inner.horizontalHeader()
    header.resizeSection(0, 70)
    header.resizeSection(1, 80)
    header.resizeSection(2, 90)
    app.processEvents()
    window.close()
    app.processEvents()
    window2 = main.MainWindow()
    widths = window2.table.get_day_column_widths()
    assert widths[:3] == [70, 80, 90]
    window2.close()
    app.quit()

