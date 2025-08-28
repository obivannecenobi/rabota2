from PySide6 import QtWidgets, QtGui

def set_text_font(name: str) -> None:
    """Set application wide text font."""
    app = QtWidgets.QApplication.instance()
    if app:
        app.setFont(QtGui.QFont(name))

def set_header_font(name: str) -> None:
    """Update fonts of all header views."""
    app = QtWidgets.QApplication.instance()
    if not app:
        return
    font = QtGui.QFont(name)
    for widget in app.allWidgets():
        if isinstance(widget, QtWidgets.QHeaderView):
            widget.setFont(font)
