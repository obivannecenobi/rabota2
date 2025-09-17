import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets  # noqa: E402

import resources  # noqa: E402

resources.register_fonts = lambda: None  # noqa: E305

import app.main as main  # noqa: E402


def test_release_dialog_style_follows_accent(monkeypatch):
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    dialog = main.ReleaseDialog(2024, 1, [], None)

    new_color = "#abcdef"
    monkeypatch.setitem(main.CONFIG, "accent_color", new_color)

    dialog.refresh_theme()

    assert new_color.lower() in dialog.table.styleSheet().lower()

    dialog.close()
    app.quit()
