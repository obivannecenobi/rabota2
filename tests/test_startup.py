import os
from PySide6 import QtWidgets

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from app import main_window


def test_startup():
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    window = main_window.MainWindow()
    assert window is not None
    window.close()
    app.quit()

