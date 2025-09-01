from PySide6 import QtWidgets


class ButtonStyleMixin:
    """Mixin providing base styling for buttons."""

    def apply_base_style(self) -> None:
        """Apply shared base QSS to the widget."""
        self.setStyleSheet(
            "border-radius:12px; padding:8px 12px; "
            "border:1px solid transparent; min-width:24px; "
            "min-height:24px; color:#e5e5e5;"
        )


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
