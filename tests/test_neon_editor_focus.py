import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets, QtCore

import resources
resources.register_fonts = lambda: None

import app.main as main
from app.effects import NeonEventFilter


def test_neon_persists_during_edit_and_stops_after():
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    main.CONFIG["neon"] = True

    table = QtWidgets.QTableWidget(1, 1)
    table.setItem(0, 0, QtWidgets.QTableWidgetItem("1"))
    table.setAttribute(QtCore.Qt.WA_Hover, True)
    table.viewport().setAttribute(QtCore.Qt.WA_Hover, True)

    filt = NeonEventFilter(table)
    table.installEventFilter(filt)
    table.viewport().installEventFilter(filt)

    table.show()
    table.setFocus()
    QtWidgets.QApplication.processEvents()

    assert getattr(table, "_neon_effect", None) is not None
    w_before = table.columnWidth(0)
    h_before = table.rowHeight(0)

    table.editItem(table.item(0, 0))
    QtWidgets.QApplication.processEvents()

    editor = table.findChild(QtWidgets.QLineEdit)
    assert editor is not None and editor.hasFocus()
    assert getattr(table, "_neon_effect", None) is not None
    assert table.columnWidth(0) == w_before
    assert table.rowHeight(0) == h_before

    other = QtWidgets.QLineEdit()
    other.setAttribute(QtCore.Qt.WA_Hover, True)
    other.show()
    other.setFocus()
    QtWidgets.QApplication.processEvents()

    assert getattr(table, "_neon_effect", None) is None
    assert table.columnWidth(0) == w_before
    assert table.rowHeight(0) == h_before

    other.close()
    table.close()
    app.quit()
