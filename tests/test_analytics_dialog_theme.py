import os
import sys
from pathlib import Path

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QT_OPENGL", "software")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

try:  # pragma: no cover - runtime import guard for CI images without OpenGL
    from PySide6 import QtGui, QtWidgets  # noqa: E402
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


def test_analytics_spin_year_border_updates_with_accent(monkeypatch):
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    dialog = main.AnalyticsDialog(2024)

    accent = "#abcdef"
    workspace = "#112233"
    monkeypatch.setitem(main.CONFIG, "accent_color", accent)
    monkeypatch.setitem(main.CONFIG, "workspace_color", workspace)

    dialog.refresh_theme()

    style = dialog.spin_year.styleSheet().lower()
    assert accent.lower() in style
    assert workspace.lower() in style

    palette = dialog.spin_year.palette()
    assert palette.color(QtGui.QPalette.Highlight).name().lower() == accent.lower()

    filt = getattr(dialog.spin_year, "_neon_filter", None)
    assert filt is not None
    assert filt._config.get("accent_color") == accent

    dialog.close()
    app.quit()
