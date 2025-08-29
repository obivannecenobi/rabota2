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
import math
from typing import Dict, Tuple


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
    if grad and len(grad) == 2:
        angle = config.get("gradient_angle", 0)
        rad = math.radians(angle)
        x2 = 0.5 + 0.5 * math.cos(rad)
        y2 = 0.5 + 0.5 * math.sin(rad)
        base = (
            f"border:1px solid #555; border-radius:8px; "
            f"background:qlineargradient(x1:0,y1:0,x2:{x2:.2f},y2:{y2:.2f}, "
            f"stop:0 {grad[0]}, stop:1 {grad[1]});"
        )
    else:
        base = (
            f"border:1px solid #555; border-radius:8px; "
            f"background-color:{workspace};"
        )
    return base, workspace


def apply_glass_effect(window: QtWidgets.QWidget, config: Dict) -> None:
    """Apply glass effect using BlurWindow based on *config*."""
    eff_type = config.get("glass_effect", "")
    opacity = float(config.get("glass_opacity", 0.5))
    blur = int(config.get("glass_blur", 10))
    try:
        from blurwindow import BlurWindow, GlobalBlur
        if eff_type:
            blur_type = getattr(GlobalBlur, eff_type, GlobalBlur.Acrylic)
            BlurWindow.blur(window.winId(), blur_type, blur, int(opacity * 255))
        else:
            BlurWindow.blur(window.winId(), GlobalBlur.CLEAR)
    except Exception:
        central = window.centralWidget()
        if eff_type:
            eff = QtWidgets.QGraphicsBlurEffect(window)
            eff.setBlurRadius(blur)
            central.setGraphicsEffect(eff)
            window.setWindowOpacity(opacity)
        else:
            central.setGraphicsEffect(None)
            window.setWindowOpacity(1.0)
