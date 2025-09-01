from PySide6 import QtWidgets, QtGui, QtCore
import re
import shiboken6
import weakref


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

    if widget is None or not shiboken6.isValid(widget):
        return

    if on:
        if getattr(widget, "_prev_effect", None) is None:
            prev = widget.graphicsEffect()
            if prev is not None and shiboken6.isValid(prev):
                try:
                    prev.setParent(None)
                except RuntimeError:
                    prev = None
            widget._prev_effect = prev
        if getattr(widget, "_neon_prev_style", None) is None:
            widget._neon_prev_style = widget.styleSheet()
        eff = QtWidgets.QGraphicsDropShadowEffect(widget)
        eff.setOffset(0, 0)
        eff.setBlurRadius(20)
        color = widget.palette().color(QtGui.QPalette.Highlight)
        eff.setColor(color)
        try:
            widget.setGraphicsEffect(eff)
        except RuntimeError:
            return
        widget.setStyleSheet(
            f"{widget._neon_prev_style}border-color:{color.name()};"
        )
        widget._neon_effect = eff
    else:
        prev = getattr(widget, "_prev_effect", None)
        widget._prev_effect = None
        try:
            if prev is not None and shiboken6.isValid(prev):
                widget.setGraphicsEffect(prev)
                try:
                    prev.setParent(widget)
                except RuntimeError:
                    pass
            else:
                widget.setGraphicsEffect(None)
        except RuntimeError:
            pass
        prev_style = getattr(widget, "_neon_prev_style", None)
        widget.setStyleSheet(prev_style or "")
        widget._neon_prev_style = None
        widget._neon_effect = None


class NeonEventFilter(QtCore.QObject):
    """Event filter toggling neon effect on hover and focus events."""

    def __init__(self, widget: QtWidgets.QWidget):
        super().__init__(widget)
        self._widget = weakref.ref(widget)
        widget.destroyed.connect(self._on_widget_destroyed)

    def _on_widget_destroyed(self, obj=None) -> None:
        widget = self._widget()
        if widget is not None and shiboken6.isValid(widget):
            widget.removeEventFilter(self)

    def _start(self) -> None:
        widget = self._widget()
        apply_neon_effect(widget, True)

    def _stop(self) -> None:
        widget = self._widget()
        if widget is None or not shiboken6.isValid(widget):
            return
        apply_neon_effect(widget, False)

    def eventFilter(self, obj, event):  # noqa: D401 - Qt event filter signature
        if event.type() in (QtCore.QEvent.HoverEnter, QtCore.QEvent.FocusIn):
            self._start()
        elif event.type() in (
            QtCore.QEvent.HoverLeave,
            QtCore.QEvent.Leave,
            QtCore.QEvent.FocusOut,
        ):
            self._stop()
        return False
