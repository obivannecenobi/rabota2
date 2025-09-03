from typing import Dict, Tuple
import math

from PySide6 import QtWidgets, QtGui, QtCore


def set_text_font(name: str = "Exo 2") -> None:
    """Set application wide text font."""
    app = QtWidgets.QApplication.instance()
    if not app:
        return
    font = QtGui.QFont(name)
    app.setFont(font)
    def _in_sidebar(w: QtWidgets.QWidget) -> bool:
        parent = w
        while parent is not None:
            if parent.objectName() == "Sidebar":
                return True
            parent = parent.parent()
        return False

    for widget in app.allWidgets():
        if _in_sidebar(widget):
            continue
        widget.setFont(font)


def set_header_font(name: str = "Exo 2") -> None:
    """Update fonts of all header views."""
    app = QtWidgets.QApplication.instance()
    if not app:
        return
    font = QtGui.QFont(name)
    for widget in app.allWidgets():
        if isinstance(widget, QtWidgets.QHeaderView):
            widget.setFont(font)


def apply_monochrome(color: QtGui.QColor) -> QtGui.QColor:
    """Convert *color* to grayscale using luminance."""
    r, g, b, _ = color.getRgb()
    luminance = int(0.299 * r + 0.587 * g + 0.114 * b)
    return QtGui.QColor(luminance, luminance, luminance)


def apply_gradient(config: Dict) -> Tuple[str, str]:
    """Return base style string and workspace color based on *config*."""
    workspace = config.get("workspace_color", "#1e1e21")
    grad = config.get("gradient_colors", None)
    if config.get("monochrome", False):
        workspace = apply_monochrome(QtGui.QColor(workspace)).name()
        if grad and len(grad) == 2:
            grad = [apply_monochrome(QtGui.QColor(c)).name() for c in grad]

    base = (
        f"border:1px solid #555; border-radius:8px; "
        f"background-color:{workspace};"
    )

    if grad and len(grad) == 2 and grad[0] != grad[1]:
        angle = config.get("gradient_angle", 0)
        rad = math.radians(angle)
        x2 = 0.5 + 0.5 * math.cos(rad)
        y2 = 0.5 + 0.5 * math.sin(rad)
        base += (
            f" background:qlineargradient(x1:0,y1:0,x2:{x2:.2f},y2:{y2:.2f}, "
            f"stop:0 {grad[0]}, stop:1 {grad[1]});"
        )

    return base, workspace
