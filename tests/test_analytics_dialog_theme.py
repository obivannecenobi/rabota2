import os
import sys
from pathlib import Path

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QT_OPENGL", "software")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

try:  # pragma: no cover - runtime import guard for CI images without OpenGL
    from PySide6 import QtWidgets  # noqa: E402
except ImportError as exc:  # pragma: no cover - skip when Qt can't load
    pytest.skip(f"PySide6 is unavailable: {exc}", allow_module_level=True)

import resources  # noqa: E402

resources.register_fonts = lambda: None  # noqa: E305

import app.main as main  # noqa: E402


def test_analytics_dialog_theme_uses_accent_and_neon(monkeypatch):
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    dialog = main.AnalyticsDialog(2024)

    new_color = "#123456"
    monkeypatch.setitem(main.CONFIG, "accent_color", new_color)

    dialog.refresh_theme()

    header_style = dialog.table.horizontalHeader().styleSheet().lower()
    assert new_color.lower() in header_style

    assert getattr(dialog.table, "_neon_effect", None) is not None
    assert dialog.table.graphicsEffect() is not None
    assert getattr(dialog.table, "_neon_filter", None) is not None

    dialog.close()
    app.quit()
