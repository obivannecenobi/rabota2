import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets

import resources
resources.register_fonts = lambda: None

import app.main as main


def _reset_day_rows(original_value):
    if original_value is None:
        main.CONFIG.pop("day_rows", None)
    else:
        main.CONFIG["day_rows"] = original_value


def test_calendar_table_uses_default_day_rows():
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    original = main.CONFIG.get("day_rows")
    main.CONFIG.pop("day_rows", None)

    try:
        table = main.ExcelCalendarTable()
        assert table.cell_tables
        inner = next(iter(table.cell_tables.values()))
        assert inner.rowCount() == main.DAY_ROWS_DEFAULT
        table.close()
    finally:
        _reset_day_rows(original)
        app.quit()


def test_settings_dialog_uses_default_day_rows():
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    original = main.CONFIG.get("day_rows")
    main.CONFIG.pop("day_rows", None)

    try:
        dlg = main.SettingsDialog()
        assert dlg.spin_day_rows.value() == main.DAY_ROWS_DEFAULT
        dlg.close()
    finally:
        _reset_day_rows(original)
        app.quit()
