from PySide6 import QtWidgets, QtGui, QtCore
from animations import NeonMotion


def set_neon(
    widget: QtWidgets.QWidget,
    color,
    intensity=255,
    pulse=False,
    motion_speed=1.0,
    mode="outer",
):
    """
    Apply neon effect to ``widget`` and optionally animate it.

    ``mode`` can be ``"outer"`` (default) for a drop shadow or ``"inner"`` for a
    colorize effect based on the widget's ``palette().buttonText()`` color.
    """

    anim = None
    motion = None

    if mode == "inner":
        eff = QtWidgets.QGraphicsColorizeEffect(widget)
        c = widget.palette().buttonText().color()
        c.setAlpha(int(intensity))
        eff.setColor(c)
        eff.setStrength(1.0)
        widget.setGraphicsEffect(eff)
        return eff, None, None

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
        duration = int(1000 / max(motion_speed, 0.1))
        anim.setDuration(duration)
        anim.setLoopCount(-1)
        anim.setEasingCurve(QtCore.QEasingCurve.InOutQuad)
        anim.start(QtCore.QAbstractAnimation.DeleteWhenStopped)

    if motion_speed and motion_speed > 0:
        motion = NeonMotion(eff, speed=motion_speed)

    return eff, anim, motion
