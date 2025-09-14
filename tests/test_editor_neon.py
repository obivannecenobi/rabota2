import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets, QtCore, QtTest, QtGui
import shiboken6

import resources
resources.register_fonts = lambda: None

import app.main as main


def test_editor_neon_clears():
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    main.CONFIG["neon"] = True
    table = main.NeonTableWidget(1, 1)
    table.setItem(0, 0, QtWidgets.QTableWidgetItem("text"))
    table.show()
    QtWidgets.QApplication.processEvents()

    index = table.model().index(0, 0)
    table.edit(index, QtWidgets.QAbstractItemView.DoubleClicked, None)
    QtWidgets.QApplication.processEvents()

    editor = table.findChild(QtWidgets.QLineEdit)
    assert editor is not None
    color = editor.palette().color(QtGui.QPalette.Highlight).name()
    assert f"border-color:{color}" in editor.styleSheet().replace(" ", "")

    table.setFocus()
    QtWidgets.QApplication.processEvents()

    assert table._active_editor is None
    if shiboken6.isValid(editor):
        assert "border-color:transparent" in editor.styleSheet().replace(" ", "")

    table.close()
    app.quit()
