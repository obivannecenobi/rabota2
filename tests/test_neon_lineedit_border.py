import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "app"))

from PySide6 import QtWidgets, QtCore, QtGui

import resources
resources.register_fonts = lambda: None

import app.main as main
from app.effects import NeonEventFilter


def test_neon_border_toggles_without_resizing():
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    main.CONFIG["neon"] = True

    edit = QtWidgets.QLineEdit()
    edit.setAttribute(QtCore.Qt.WA_Hover, True)
    edit.setStyleSheet("border:1px solid transparent;")
    filt = NeonEventFilter(edit)
    edit.installEventFilter(filt)
    edit.show()
    edit.resize(100, 30)
    QtWidgets.QApplication.processEvents()

    size_before = edit.size()
    edit.setFocus()
    QtWidgets.QApplication.processEvents()
    color = edit.palette().color(QtGui.QPalette.Highlight).name()
    assert f"border-color:{color}" in edit.styleSheet().replace(" ", "")
    assert edit.size() == size_before

    other = QtWidgets.QLineEdit()
    other.setAttribute(QtCore.Qt.WA_Hover, True)
    other.setStyleSheet("border:1px solid transparent;")
    other.installEventFilter(NeonEventFilter(other))
    other.show()
    other.setFocus()
    QtWidgets.QApplication.processEvents()

    assert "border-color:transparent" in edit.styleSheet().replace(" ", "")
    assert edit.size() == size_before

    other.close()
    edit.close()
    app.quit()
