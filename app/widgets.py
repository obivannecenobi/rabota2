from PySide6 import QtWidgets, QtGui, QtCore

from effects import apply_neon_effect, neon_enabled
from resources import icon


class ButtonStyleMixin:
    """Mixin providing sidebar-style appearance and hover behaviour."""

    _base_style = (
        "border-radius:12px; padding:8px 12px; "
        "border:1px solid transparent; min-width:24px; "
        "min-height:24px; color:#e5e5e5; background:transparent;"
    )

    def apply_base_style(self) -> None:
        """Apply the base style and enable hover tracking."""
        self.setStyleSheet(self._base_style)
        self.setCursor(QtCore.Qt.PointingHandCursor)
        self.setAttribute(QtCore.Qt.WA_Hover, True)

    # --- helpers -----------------------------------------------------
    def _accent_color(self) -> str:
        return self.palette().color(QtGui.QPalette.Highlight).name()

    def _hover_style(self) -> str:
        try:
            from . import main  # type: ignore

            thickness = main.CONFIG.get("neon_thickness", 1)
        except Exception:  # pragma: no cover
            thickness = 1
        return (
            f"border:{thickness}px solid {self._accent_color()}; "
            "background-color:rgba(255,255,255,0.08);"
        )

    def _apply_hover(self, on: bool) -> None:
        self.setStyleSheet(self._base_style + (self._hover_style() if on else ""))

    # --- events ------------------------------------------------------
    def enterEvent(self, event):  # noqa: D401
        self._apply_hover(True)
        if neon_enabled():
            apply_neon_effect(self, True)
        super().enterEvent(event)

    def leaveEvent(self, event):  # noqa: D401
        selected = bool(self.property("neon_selected"))
        self._apply_hover(selected)
        if neon_enabled():
            apply_neon_effect(self, selected)
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


class SpinButton(StyledToolButton):
    """Small tool button used to step a spin box."""

    def __init__(self, spinbox: QtWidgets.QAbstractSpinBox, step: int, parent=None):
        super().__init__(parent)
        self._spinbox = spinbox
        self._step = 1 if step >= 0 else -1
        icon_name = "chevron-up" if self._step > 0 else "chevron-down"
        self.setIcon(icon(icon_name))
        self.setIconSize(QtCore.QSize(12, 12))
        self.setFixedSize(20, 20)
        self.setAutoRepeat(True)
        self.clicked.connect(self._on_clicked)

    def _on_clicked(self) -> None:
        if self._step > 0:
            self._spinbox.stepUp()
        else:
            self._spinbox.stepDown()
