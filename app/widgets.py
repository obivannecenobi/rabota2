import math

from PySide6 import QtWidgets, QtGui, QtCore

from effects import apply_neon_effect, neon_enabled
from resources import icon


class ButtonStyleMixin:
    """Mixin providing sidebar-style appearance and hover behaviour."""

    def _base_style(self) -> str:
        """Build the base style using current configuration."""
        try:
            from . import main  # type: ignore

            thickness = main.CONFIG.get("neon_thickness", 1)
            grad = main.CONFIG.get("gradient_colors", ["#39ff14", "#2d7cdb"])
            angle = main.CONFIG.get("gradient_angle", 0)
        except Exception:  # pragma: no cover
            thickness = 1
            grad = ["#39ff14", "#2d7cdb"]
            angle = 0
        rad = math.radians(angle)
        x2 = 0.5 + 0.5 * math.cos(rad)
        y2 = 0.5 + 0.5 * math.sin(rad)
        return (
            "border-radius:16px; padding:8px 12px; "
            f"border:{thickness}px solid transparent; min-width:24px; min-height:24px; "
            "color:white; "
            f"background: qlineargradient(x1:0,y1:0,x2:{x2:.2f},y2:{y2:.2f}, "
            f"stop:0 {grad[0]}, stop:1 {grad[1]});"
        )

    def apply_base_style(self) -> None:
        """Apply the base style and enable hover tracking."""
        self.setStyleSheet(self._base_style())
        self.setCursor(QtCore.Qt.PointingHandCursor)
        self.setAttribute(QtCore.Qt.WA_Hover, True)

    # --- helpers -----------------------------------------------------
    def _accent_color(self) -> str:
        return self.palette().color(QtGui.QPalette.Highlight).name()

    def _hover_style(self) -> str:
        return (
            f"border-color:{self._accent_color()}; "
            f"color:{self._accent_color()};"
        )

    def _apply_hover(self, on: bool) -> None:
        if on:
            self.setStyleSheet(self._base_style() + self._hover_style())
        else:
            self.setStyleSheet(self._base_style())

    # --- events ------------------------------------------------------
    def enterEvent(self, event):  # noqa: D401
        self._apply_hover(True)
        if neon_enabled():
            apply_neon_effect(self, True)
        super().enterEvent(event)

    def leaveEvent(self, event):  # noqa: D401
        selected = bool(self.property("neon_selected"))
        if neon_enabled():
            apply_neon_effect(self, selected)
        self._apply_hover(selected)
        super().leaveEvent(event)


class StyledPushButton(ButtonStyleMixin, QtWidgets.QPushButton):
    """QPushButton with shared styling mixin."""

    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        self.apply_base_style()


class StyledToolButton(ButtonStyleMixin, QtWidgets.QToolButton):
    """QToolButton with shared styling mixin."""

    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        self.apply_base_style()
