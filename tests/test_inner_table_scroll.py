import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets, QtCore, QtGui

import resources
resources.register_fonts = lambda: None

import app.main as main


def test_inner_table_scrolling_works(monkeypatch):
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    monkeypatch.setattr(main.ExcelCalendarTable, "load_month_data", lambda self, y, m: None)

    table = main.ExcelCalendarTable()
    inner = table._create_inner_table()
    inner.setRowCount(30)
    for r in range(30):
        inner.setItem(r, 0, QtWidgets.QTableWidgetItem(str(r)))
    inner.resize(100, 100)
    inner.show()

    assert inner.verticalScrollBarPolicy() == QtCore.Qt.ScrollBarAlwaysOff
    assert inner.horizontalScrollBarPolicy() == QtCore.Qt.ScrollBarAlwaysOff

    vbar = inner.verticalScrollBar()
    start = vbar.value()

    wheel_event = QtGui.QWheelEvent(
        QtCore.QPointF(50, 50),
        QtCore.QPointF(50, 50),
        QtCore.QPoint(),
        QtCore.QPoint(0, -120),
        QtCore.Qt.NoButton,
        QtCore.Qt.NoModifier,
        QtCore.Qt.ScrollUpdate,
        False,
    )
    QtWidgets.QApplication.sendEvent(inner.viewport(), wheel_event)
    QtWidgets.QApplication.processEvents()
    assert vbar.value() != start

    vbar.setValue(0)
    QtWidgets.QApplication.processEvents()
    key_event = QtGui.QKeyEvent(
        QtCore.QEvent.KeyPress, QtCore.Qt.Key_PageDown, QtCore.Qt.NoModifier
    )
    QtWidgets.QApplication.sendEvent(vbar, key_event)
    QtWidgets.QApplication.processEvents()
    assert vbar.value() > 0

    inner.close()
    app.quit()
