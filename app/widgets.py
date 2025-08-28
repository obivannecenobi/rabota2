from PySide6 import QtWidgets, QtCore, QtGui


class ButtonStyleMixin(object):
    """Mixin providing base styling and neon effect helpers for buttons."""

    def apply_base_style(self):
        """Apply shared base QSS to the widget."""
        self.setStyleSheet(
            "border-radius:12px; padding:8px 12px; border:1px solid transparent;"
        )

    def apply_neon_color(self, color, intensity=255, pulse=False):
        """Apply neon-like color glow using drop shadow effect.

        Parameters
        ----------
        color : QtGui.QColor or str
            Color of the neon glow.
        intensity : int
            Alpha channel value for the glow.
        pulse : bool
            If True, animate the blur radius to create a pulsing effect.
        """
        eff = QtWidgets.QGraphicsDropShadowEffect(self)
        eff.setOffset(0, 0)
        eff.setBlurRadius(20)
        c = QtGui.QColor(color)
        c.setAlpha(intensity)
        eff.setColor(c)
        self.setGraphicsEffect(eff)
        if pulse:
            anim = QtCore.QPropertyAnimation(eff, b"blurRadius", self)
            anim.setStartValue(20)
            anim.setEndValue(40)
            anim.setDuration(1000)
            anim.setLoopCount(-1)
            anim.setEasingCurve(QtCore.QEasingCurve.InOutQuad)
            anim.start(QtCore.QAbstractAnimation.DeleteWhenStopped)
            self._neon_anim = anim
        else:
            self._neon_anim = None
        return eff


class StyledPushButton(ButtonStyleMixin, QtWidgets.QPushButton):
    """QPushButton with shared styling mixin."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.apply_base_style()


class StyledToolButton(ButtonStyleMixin, QtWidgets.QToolButton):
    """QToolButton with shared styling mixin."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.apply_base_style()
