import os
import sys
from pathlib import Path

from PySide6 import QtWidgets

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

import resources  # noqa: E402

resources.register_fonts = lambda: None

import app.main as main  # noqa: E402


def test_month_data_respects_updated_save_path(tmp_path, monkeypatch):
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path / "config"))
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    original_base = main.BASE_SAVE_PATH
    original_save_path = main.CONFIG.get("save_path")
    table = None
    try:
        old_dir = tmp_path / "old"
        new_dir = tmp_path / "new"
        main.BASE_SAVE_PATH = str(old_dir)
        main.CONFIG["save_path"] = main.BASE_SAVE_PATH

        table = main.ExcelCalendarTable()
        table.save_current_month()

        filename = f"{table.year:04d}-{table.month:02d}.json"
        old_file = old_dir / "months" / filename
        assert old_file.exists()

        old_file.unlink()
        assert not old_file.exists()

        main.CONFIG["save_path"] = str(new_dir)
        main.BASE_SAVE_PATH = str(new_dir)

        table.save_current_month()

        new_file = new_dir / "months" / filename
        assert new_file.exists()
        assert not old_file.exists()
    finally:
        if table is not None:
            table.deleteLater()
        app.processEvents()
        app.quit()
        main.BASE_SAVE_PATH = original_base
        if original_save_path is not None:
            main.CONFIG["save_path"] = original_save_path
        else:
            main.CONFIG.pop("save_path", None)

