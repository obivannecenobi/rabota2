from PySide6 import QtCore, QtWidgets
import math

class NeonMotion(QtCore.QObject):
    """Timer driven animation for cyclic neon offset movement."""

    def __init__(self, effect: QtWidgets.QGraphicsEffect, radius: float = 2.0, speed: float = 1.0):
        super().__init__(effect)
        self._effect = effect
        self._radius = radius
        self._angle = 0.0
        self._timer = QtCore.QTimer(self)
        interval = max(10, int(30 / max(speed, 0.1)))
        self._timer.timeout.connect(self._update)
        self._timer.start(interval)

    def _update(self) -> None:
        self._angle = (self._angle + 5) % 360
        x = self._radius * math.cos(math.radians(self._angle))
        y = self._radius * math.sin(math.radians(self._angle))
        self._effect.setOffset(x, y)

    def stop(self) -> None:
        self._timer.stop()
        self._effect.setOffset(0, 0)
