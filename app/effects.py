from PySide6 import QtWidgets, QtGui, QtCore
import shiboken6
import weakref


class FixedDropShadowEffect(QtWidgets.QGraphicsDropShadowEffect):
    """Drop shadow effect that preserves the original bounding rectangle."""

    def boundingRectFor(  # noqa: D401 - Qt override signature
        self, rect: QtCore.QRectF
    ) -> QtCore.QRectF:
        """Return the original ``rect`` without expanding it."""

        return rect


def neon_enabled(config: dict) -> bool:
    """Return ``True`` if neon highlighting is enabled in *config*."""
    return bool(config.get("neon", False))


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

    eff = FixedDropShadowEffect(widget)
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


def apply_neon_effect(
    widget: QtWidgets.QWidget,
    on: bool = True,
    shadow: bool = True,
    border: bool = True,
    *,
    config: dict | None = None,
) -> None:
    """Toggle neon highlight effect on *widget*.

    Parameters
    ----------
    widget: QtWidgets.QWidget
        Target widget.
    on: bool
        Whether to enable the effect.  When ``False`` the previous style and
        effect (if any) are restored.
    shadow: bool
        If ``True`` (default) a :class:`~PySide6.QtWidgets.QGraphicsDropShadowEffect`
        is applied.  Set to ``False`` to skip the drop shadow, keeping only the
        color adjustments.  This is useful for widgets like ``QLabel`` where the
        shadow is visually undesirable.
    border: bool
        When ``True`` (default) a colored border matching the highlight color is
        added to the widget.  Set to ``False`` to leave the widget border
        unchanged.
    """

    if widget is None or not shiboken6.isValid(widget):
        return

    try:
        thickness = int(config.get("neon_thickness", 1)) if config else 1
    except (TypeError, ValueError):
        thickness = 1
    thickness = max(0, thickness)

    if on:
        if getattr(widget, "_neon_prev_effect", None) is None:
            prev = widget.graphicsEffect()
            if prev is not None and shiboken6.isValid(prev):
                try:
                    prev.setParent(None)
                except RuntimeError:
                    prev = None
            widget._neon_prev_effect = prev
        if getattr(widget, "_neon_prev_style", None) is None:
            widget._neon_prev_style = widget.styleSheet()
        prev_style = widget._neon_prev_style or ""
        color = widget.palette().color(QtGui.QPalette.Highlight)
        eff = None
        blur_radius = 20
        intensity = 255
        if config:
            try:
                blur_radius = int(config.get("neon_size", blur_radius))
            except (TypeError, ValueError):
                blur_radius = 20
            try:
                intensity = int(config.get("neon_intensity", intensity))
            except (TypeError, ValueError):
                intensity = 255
        blur_radius = max(0, blur_radius)
        intensity = max(0, min(255, intensity))

        if shadow:
            eff = FixedDropShadowEffect(widget)
            eff.setOffset(0, 0)
            eff.setBlurRadius(blur_radius)
            effect_color = QtGui.QColor(color)
            effect_color.setAlpha(intensity)
            eff.setColor(effect_color)
            try:
                widget.setGraphicsEffect(eff)
            except RuntimeError:
                return
        else:
            widget.setGraphicsEffect(None)
        if border:
            border_style = f" border:{thickness}px solid {color.name()};"
        else:
            border_style = ""
        text_style = f" color:{color.name()};"
        widget.setStyleSheet(prev_style + text_style + border_style)
        widget._neon_effect = eff
    else:
        prev = getattr(widget, "_neon_prev_effect", None)
        widget._neon_prev_effect = None
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
        if prev_style:
            widget.setStyleSheet(prev_style)
        elif isinstance(widget, QtWidgets.QLabel):
            widget.setStyleSheet("")
        else:
            widget.setStyleSheet("border:0px solid transparent;")
        widget._neon_prev_style = None
        widget._neon_effect = None


class NeonEventFilter(QtCore.QObject):
    """Event filter toggling neon effect on hover and focus events."""

    def __init__(self, widget: QtWidgets.QWidget, config: dict):
        super().__init__(widget)
        self._widget = weakref.ref(widget)
        self._config = config
        widget.destroyed.connect(self._on_widget_destroyed)

    def _on_widget_destroyed(self, obj=None) -> None:
        widget = self._widget()
        if widget is not None and shiboken6.isValid(widget):
            widget.removeEventFilter(self)

    def _start(self) -> None:
        if not neon_enabled(self._config):
            return
        widget = self._widget()
        if getattr(widget, "_neon_effect", None):
            return
        apply_neon_effect(widget, True, config=self._config)

    def _stop(self) -> None:
        if not neon_enabled(self._config):
            return
        widget = self._widget()
        if widget is None or not shiboken6.isValid(widget):
            return
        apply_neon_effect(widget, False, config=self._config)

    def eventFilter(self, obj, event):  # noqa: D401 - Qt event filter signature
        if not neon_enabled(self._config):
            return False

        widget = self._widget()
        if widget is None or not shiboken6.isValid(widget):
            return False

        etype = event.type()

        if etype in (
            QtCore.QEvent.FocusIn,
            QtCore.QEvent.MouseButtonDblClick,
        ):
            self._start()
            if etype == QtCore.QEvent.FocusIn:
                focus = QtWidgets.QApplication.focusWidget()
                if (
                    focus is not None
                    and focus is not widget
                    and widget.isAncestorOf(focus)
                ):
                    focus.installEventFilter(self)

        elif etype == QtCore.QEvent.FocusOut:
            focus = QtWidgets.QApplication.focusWidget()
            if obj is widget and focus is not None and widget.isAncestorOf(focus):
                focus.installEventFilter(self)
            else:
                self._stop()

        elif etype in (
            QtCore.QEvent.HoverLeave,
            QtCore.QEvent.Leave,
        ):
            if not widget.hasFocus():
                self._stop()

        return False


def update_neon_filters(root: QtWidgets.QWidget, config: dict) -> None:
    """Recursively update neon event filters according to configuration."""

    if root is None or not shiboken6.isValid(root):
        return

    enabled = neon_enabled(config)

    def _process(widget: QtWidgets.QWidget) -> None:
        if widget is None or not shiboken6.isValid(widget):
            return
        if not widget.testAttribute(QtCore.Qt.WA_Hover):
            return
        filt = getattr(widget, "_neon_filter", None)
        if enabled and filt is None:
            filt = NeonEventFilter(widget, config)
            widget.installEventFilter(filt)
            if hasattr(widget, "viewport"):
                widget.viewport().installEventFilter(filt)
            widget._neon_filter = filt
        elif enabled and filt is not None:
            filt._config = config
        elif not enabled and filt is not None:
            try:
                widget.removeEventFilter(filt)
                if hasattr(widget, "viewport"):
                    widget.viewport().removeEventFilter(filt)
            except RuntimeError:
                pass
            apply_neon_effect(widget, False, config=config)
            widget._neon_filter = None

    widgets = [root] + root.findChildren(QtWidgets.QWidget)
    for w in widgets:
        _process(w)
