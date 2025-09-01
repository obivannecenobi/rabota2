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
    for widget in app.allWidgets():
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
    """Apply glass effect to the main workspace."""
    enabled = config.get("glass_enabled", False)
    eff_type = config.get("glass_effect") or "Acrylic"
    opacity = float(config.get("glass_opacity", 0.5))
    blur = int(config.get("glass_blur", 10))

    central = window.centralWidget()
    if not central:
        return

    # remove fallback layer if present
    back = getattr(window, "_glass_back", None)

    try:
        from blurwindow import BlurWindow, GlobalBlur
        if enabled and getattr(BlurWindow, "is_supported", lambda: True)():
            if back:
                back.setParent(None)
                window._glass_back = None
            blur_type = getattr(GlobalBlur, eff_type, getattr(GlobalBlur, "Acrylic", None))
            BlurWindow.blur(central.winId(), blur_type, blur, int(opacity * 255))
        else:
            try:
                BlurWindow.blur(central.winId(), getattr(GlobalBlur, "CLEAR", 0))
            except Exception:
                pass
            enabled = False
            raise Exception
    except Exception:
        if enabled:
            if back is None:
                back = QtWidgets.QWidget(central)
                back.setObjectName("_glass_back")
                back.setAttribute(QtCore.Qt.WA_TransparentForMouseEvents, True)
                window._glass_back = back
            back.setGeometry(central.rect())
            eff = QtWidgets.QGraphicsBlurEffect(back)
            eff.setBlurRadius(blur)
            back.setGraphicsEffect(eff)
            color = QtGui.QColor(config.get("workspace_color", "#1e1e21"))
            color.setAlpha(int(opacity * 255))
            pal = back.palette()
            pal.setColor(back.backgroundRole(), color)
            back.setAutoFillBackground(True)
            back.setPalette(pal)
            back.lower()
            back.show()
            central.raise_()
        elif back is not None:
            back.setParent(None)
            window._glass_back = None
