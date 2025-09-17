import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets

import resources
resources.register_fonts = lambda: None

import app.main as main
from widgets import StyledPushButton


def _get_app():
    return QtWidgets.QApplication.instance() or QtWidgets.QApplication([])


def test_plus_minus_icons_available_for_themes():
    app = _get_app()
    try:
        for theme in ("dark", "light"):
            resources.load_icons(theme)
            for name in ("plus", "minus"):
                icon_obj = resources.icon(name)
                assert not icon_obj.isNull(), f"{theme}:{name}"
                assert name in resources.ICONS
    finally:
        app.quit()


def test_release_dialog_buttons_receive_icons():
    app = _get_app()
    try:
        resources.load_icons("dark")
        dialog = main.ReleaseDialog(2024, 1, [])
        buttons = {btn.text(): btn for btn in dialog.findChildren(StyledPushButton)}
        expected = {
            "Добавить запись": "plus",
            "Удалить запись": "minus",
            "Сохранить": "save",
            "Закрыть": "x",
        }
        for text, icon_name in expected.items():
            assert text in buttons
            btn = buttons[text]
            assert not btn.icon().isNull(), f"Button '{text}' is missing icon"
            if text in ("Добавить запись", "Удалить запись"):
                # Ensure dialog pulls themed icons loaded earlier.
                assert resources.icon(icon_name).cacheKey() == btn.icon().cacheKey()
        dialog.close()
    finally:
        app.quit()
