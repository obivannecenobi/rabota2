from PySide6 import QtWidgets, QtGui


def set_neon(
    widget: QtWidgets.QWidget,
    color,
    intensity=255,
    mode="outer",
):
    """
    Apply neon effect to ``widget``.

    ``mode`` can be ``"outer"`` (default) for a drop shadow or ``"inner"`` for a
    colorize effect based on the widget's ``palette().buttonText()`` color.
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

    return eff
