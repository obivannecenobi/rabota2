from PySide6 import QtWidgets, QtGui, QtCore
import re
import shiboken6


def set_neon(
    widget: QtWidgets.QWidget,
    color,
    intensity=255,
    mode="outer",
    pulse=False,
):
    """Apply neon effect to ``widget``.

    ``mode`` can be ``"outer"`` (default) for a drop shadow or ``"inner"" for
    a colorize effect based on the widget's ``palette().buttonText()`` color.
    """

    if mode == "inner":
        eff = QtWidgets.QGraphicsColorizeEffect(widget)
        c = widget.palette().buttonText().color()
        c.setAlpha(int(intensity))
        eff.setColor(c)
        eff.setStrength(1.0)
        widget.setGraphicsEffect(eff)
        return eff

    eff = QtWidgets.QGraphicsDropShadowEffect(widget)
    eff.setOffset(0, 0)
    eff.setBlurRadius(20)
    c = QtGui.QColor(color)
    c.setAlpha(int(intensity))
    eff.setColor(c)
    widget.setGraphicsEffect(eff)

    if pulse:
        anim = QtCore.QPropertyAnimation(eff, b"blurRadius", widget)
        anim.setStartValue(20)
        anim.setEndValue(40)
        anim.setDuration(1000)
        anim.setLoopCount(-1)
        anim.setEasingCurve(QtCore.QEasingCurve.InOutQuad)
        anim.start(QtCore.QAbstractAnimation.DeleteWhenStopped)
        eff._neon_anim = anim

    return eff


def apply_neon_effect(widget: QtWidgets.QWidget, on: bool = True) -> None:
    """Toggle neon highlight effect on *widget*.

    When ``on`` is ``True`` a drop shadow with the palette's highlight color is
    applied. Passing ``False`` restores the previous effect, if any.
    """

    if on:
        if widget.property("_prev_effect") is None:
            widget.setProperty("_prev_effect", widget.graphicsEffect())
        if getattr(widget, "_neon_prev_style", None) is None:
            widget._neon_prev_style = widget.styleSheet()
            match = re.search(r"border-radius:\s*([0-9.]+)px", widget._neon_prev_style)
            widget._neon_prev_radius = match.group(1) if match else None
        eff = QtWidgets.QGraphicsDropShadowEffect(widget)
        eff.setOffset(0, 0)
        eff.setBlurRadius(20)
        color = widget.palette().color(QtGui.QPalette.Highlight)
        eff.setColor(color)
        widget.setGraphicsEffect(eff)
        parts = [widget._neon_prev_style, f"border:1px solid {color.name()};"]
        if widget._neon_prev_radius is not None:
            parts.append(f"border-radius:{widget._neon_prev_radius}px;")
        widget.setStyleSheet("".join(parts))
        widget._neon_effect = eff
    else:
        prev = widget.property("_prev_effect")
        widget.setProperty("_prev_effect", None)
        if prev is not None and shiboken6.isValid(prev):
            widget.setGraphicsEffect(prev)
        else:
            widget.setGraphicsEffect(None)
        prev_style = getattr(widget, "_neon_prev_style", None)
        if prev_style is not None:
            widget.setStyleSheet(prev_style)
        widget._neon_prev_style = None
        widget._neon_prev_radius = None
        widget._neon_effect = None


class NeonEventFilter(QtCore.QObject):
    """Event filter toggling neon effect on hover and focus events."""

    def __init__(self, widget: QtWidgets.QWidget):
        super().__init__(widget)
        self._widget = widget
        widget.destroyed.connect(lambda: widget.removeEventFilter(self))

    def eventFilter(self, obj, event):  # noqa: D401 - Qt event filter signature
        if event.type() in (QtCore.QEvent.HoverEnter, QtCore.QEvent.FocusIn):
            apply_neon_effect(self._widget, True)
        elif event.type() in (
            QtCore.QEvent.HoverLeave,
            QtCore.QEvent.Leave,
            QtCore.QEvent.FocusOut,
        ):
            # clear stored effect and disable neon
            _ = self._widget.property("_prev_effect")
            self._widget.setProperty("_prev_effect", None)
            self._widget.setGraphicsEffect(None)
            apply_neon_effect(self._widget, False)
        return False
