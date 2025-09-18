import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets, QtGui

import resources

resources.register_fonts = lambda: None

import app.main as main
from app.effects import FixedDropShadowEffect


def test_day_container_neon_activation():
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    table = main.ExcelCalendarTable()
    coords = next(
        (coord for coord, dt in table.date_map.items() if dt.month == table.month),
        None,
    )
    assert coords is not None

    container = table.cell_containers[coords]
    base_style = container.styleSheet()

    table._set_active_day(coords, transient=False)
    QtWidgets.QApplication.processEvents()

    effect = container.graphicsEffect()
    assert isinstance(effect, FixedDropShadowEffect)

    highlight_color = container.palette().color(QtGui.QPalette.Highlight).name().lower()
    assert highlight_color in container.styleSheet().lower()

    table._clear_active_day(coords, transient=False)
    QtWidgets.QApplication.processEvents()

    assert container.graphicsEffect() is None
    assert container.styleSheet() == base_style

    table.close()
    app.quit()
