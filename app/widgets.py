import math

from PySide6 import QtWidgets, QtGui, QtCore

from effects import apply_neon_effect
from resources import icon
from config import CONFIG


class ButtonStyleMixin:
    """Mixin providing sidebar-style appearance and hover behaviour."""

    def __init__(
        self,
        *args,
        gradient_colors=None,
        gradient_angle=0,
        neon_thickness=1,
        neon_idle_intensity=0.45,
        neon_idle_thickness=0.6,
        neon_hover_intensity=1.0,
        neon_hover_thickness=1.0,
        **kwargs,
    ) -> None:
        super().__init__(*args, **kwargs)
        self._gradient_colors = gradient_colors or ["#39ff14", "#2d7cdb"]
        self._gradient_angle = gradient_angle
        self._neon_thickness = neon_thickness
        self._state_styles: dict[str, str] = {}
        self._current_state = "idle"
        self._neon_profiles = {
            "idle": {
                "intensity_scale": float(max(0.0, neon_idle_intensity)),
                "thickness_scale": float(max(0.0, neon_idle_thickness)),
            },
            "hover": {
                "intensity_scale": float(max(0.0, neon_hover_intensity)),
                "thickness_scale": float(max(0.0, neon_hover_thickness)),
            },
        }
        self._update_state_styles()

    def update_gradient(
        self,
        gradient_colors=None,
        gradient_angle=None,
        neon_thickness=None,
    ) -> None:
        """Update stored gradient configuration."""
        if gradient_colors is not None:
            self._gradient_colors = gradient_colors
        if gradient_angle is not None:
            self._gradient_angle = gradient_angle
        if neon_thickness is not None:
            self._neon_thickness = neon_thickness
        self._update_state_styles()

    def _base_style(self) -> str:
        """Build the base style using current configuration."""
        thickness = self._neon_thickness
        grad = self._gradient_colors
        angle = self._gradient_angle
        rad = math.radians(angle)
        x2 = 0.5 + 0.5 * math.cos(rad)
        y2 = 0.5 + 0.5 * math.sin(rad)
        pad_v = max(0, 8 - thickness)
        pad_h = max(0, 12 - thickness)
        return (
            "border-radius:16px;"
            f"padding:{pad_v}px {pad_h}px;"
            f"border:{thickness}px solid transparent; min-width:24px; min-height:24px;"
            "color:white;"
            f"background: qlineargradient(x1:0,y1:0,x2:{x2:.2f},y2:{y2:.2f},"
            f"stop:0 {grad[0]}, stop:1 {grad[1]});"
        )

    def apply_base_style(self) -> None:
        """Apply the base style, enable hover tracking and neon."""
        self.setCursor(QtCore.Qt.PointingHandCursor)
        self.setAttribute(QtCore.Qt.WA_Hover, True)
        self.apply_neon_state("idle")

    # --- helpers -----------------------------------------------------
    def _accent_color(self) -> str:
        return self.palette().color(QtGui.QPalette.Highlight).name()

    @staticmethod
    def _blend_colors(base: str, accent: str, ratio: float) -> str:
        """Blend ``base`` colour with ``accent`` using the provided ratio."""

        base_color = QtGui.QColor(base)
        accent_color = QtGui.QColor(accent)
        if not base_color.isValid() or not accent_color.isValid():
            return base
        ratio = max(0.0, min(1.0, ratio))
        red = round(base_color.red() * (1 - ratio) + accent_color.red() * ratio)
        green = round(base_color.green() * (1 - ratio) + accent_color.green() * ratio)
        blue = round(base_color.blue() * (1 - ratio) + accent_color.blue() * ratio)
        return QtGui.QColor(red, green, blue).name()

    def _hover_style(self) -> str:
        accent = self._accent_color()
        blended_colors = [
            self._blend_colors(color, accent, 0.25) for color in self._gradient_colors
        ]
        angle = math.radians(self._gradient_angle)
        x2 = 0.5 + 0.5 * math.cos(angle)
        y2 = 0.5 + 0.5 * math.sin(angle)
        gradient_stops = []
        if blended_colors:
            gradient_stops.append(f"stop:0 {blended_colors[0]}")
        gradient_stops.append(f"stop:0.5 {accent}")
        if blended_colors:
            gradient_stops.append(f"stop:1 {blended_colors[-1]}")
        gradient_definition = ", ".join(gradient_stops)
        return (
            "border-radius:16px;"
            f"background: qlineargradient(x1:0,y1:0,x2:{x2:.2f},y2:{y2:.2f},{gradient_definition});"
            f"border:{self._neon_thickness}px solid {accent};"
            f"color:{accent};"
        )

    def _apply_hover(self, on: bool) -> None:
        state = "hover" if on else "idle"
        self._apply_style_state(state)

    def _update_state_styles(self) -> None:
        base = self._base_style()
        self._state_styles = {
            "idle": base,
            "hover": base + self._hover_style(),
        }

    def _apply_style_state(self, state: str) -> None:
        if state not in self._state_styles:
            self._update_state_styles()
        self._current_state = state
        self.setStyleSheet(self._state_styles[state])

    def _apply_neon_profile(self, state: str) -> None:
        profile = self._neon_profiles.get(state, self._neon_profiles["idle"])
        intensity = profile.get("intensity_scale", 1.0)
        thickness = profile.get("thickness_scale", 1.0)
        apply_neon_effect(
            self,
            True,
            intensity_scale=intensity,
            thickness_scale=thickness,
            config=CONFIG,
        )

    def apply_neon_state(self, state: str | bool) -> None:
        """Apply the requested neon state and reset cached styles."""

        state_name = "hover" if state in ("hover", True) else "idle"
        self._apply_style_state(state_name)
        # Reset cached stylesheet to avoid repeated CSS concatenation when the
        # effect is reapplied (e.g. after theme changes).
        self._neon_prev_style = None
        self._apply_neon_profile(state_name)

    # --- events ------------------------------------------------------
    def enterEvent(self, event):  # noqa: D401
        self.apply_neon_state("hover")
        super().enterEvent(event)

    def leaveEvent(self, event):  # noqa: D401
        selected = bool(self.property("neon_selected"))
        state = "hover" if selected else "idle"
        self.apply_neon_state(state)
        super().leaveEvent(event)


class StyledPushButton(ButtonStyleMixin, QtWidgets.QPushButton):
    """QPushButton with shared styling mixin."""

    def __init__(
        self,
        *args,
        gradient_colors=None,
        gradient_angle=0,
        neon_thickness=1,
        **kwargs,
    ) -> None:
        super().__init__(
            *args,
            gradient_colors=gradient_colors,
            gradient_angle=gradient_angle,
            neon_thickness=neon_thickness,
            **kwargs,
        )
        self.apply_base_style()


class StyledToolButton(ButtonStyleMixin, QtWidgets.QToolButton):
    """QToolButton with shared styling mixin and centered content."""

    def __init__(
        self,
        *args,
        gradient_colors=None,
        gradient_angle=0,
        neon_thickness=1,
        **kwargs,
    ) -> None:
        super().__init__(
            *args,
            gradient_colors=gradient_colors,
            gradient_angle=gradient_angle,
            neon_thickness=neon_thickness,
            **kwargs,
        )
        self._content_spacing = 8
        self.apply_base_style()

    def setContentSpacing(self, spacing: int) -> None:
        """Update spacing between icon and text when both are visible."""

        self._content_spacing = max(0, int(spacing))
        self.update()

    def contentSpacing(self) -> int:
        """Return spacing between icon and text."""

        return self._content_spacing

    def paintEvent(self, event):  # noqa: D401
        option = QtWidgets.QStyleOptionToolButton()
        self.initStyleOption(option)

        painter = QtWidgets.QStylePainter(self)
        painter.setRenderHint(QtGui.QPainter.Antialiasing, True)
        base_option = QtWidgets.QStyleOptionToolButton(option)
        base_option.icon = QtGui.QIcon()
        base_option.text = ""
        painter.drawComplexControl(QtWidgets.QStyle.CC_ToolButton, base_option)

        contents = self.style().subControlRect(
            QtWidgets.QStyle.CC_ToolButton,
            option,
            QtWidgets.QStyle.SC_ToolButton,
            self,
        )

        icon = option.icon
        text = option.text
        style = option.toolButtonStyle
        spacing = self.style().pixelMetric(
            QtWidgets.QStyle.PM_ToolBarItemSpacing, option, self
        )
        if spacing < 0:
            spacing = 0
        spacing = max(spacing, self._content_spacing if text and not icon.isNull() else 0)

        icon_size = option.iconSize
        icon_rect = QtCore.QRect()
        text_rect = QtCore.QRect()

        if style == QtCore.Qt.ToolButtonIconOnly or not text:
            if not icon.isNull():
                icon_rect = QtCore.QRect(
                    contents.x()
                    + max(0, (contents.width() - icon_size.width()) // 2),
                    contents.y()
                    + max(0, (contents.height() - icon_size.height()) // 2),
                    icon_size.width(),
                    icon_size.height(),
                )
        elif icon.isNull():
            text_rect = contents
        elif style == QtCore.Qt.ToolButtonTextBesideIcon:
            metrics = self.fontMetrics()
            text_width = metrics.size(QtCore.Qt.TextShowMnemonic, text).width()
            total_width = icon_size.width() + spacing + text_width
            start_x = contents.x() + max(0, (contents.width() - total_width) // 2)
            icon_rect = QtCore.QRect(
                start_x,
                contents.y()
                + max(0, (contents.height() - icon_size.height()) // 2),
                icon_size.width(),
                icon_size.height(),
            )
            available_width = contents.right() - (icon_rect.right() + spacing) + 1
            text_rect = QtCore.QRect(
                icon_rect.right() + spacing,
                contents.y(),
                max(0, min(text_width, available_width)),
                contents.height(),
            )
        elif style == QtCore.Qt.ToolButtonTextUnderIcon:
            metrics = self.fontMetrics()
            text_height = metrics.size(QtCore.Qt.TextShowMnemonic, text).height()
            total_height = icon_size.height() + spacing + text_height
            start_y = contents.y() + max(0, (contents.height() - total_height) // 2)
            icon_rect = QtCore.QRect(
                contents.x()
                + max(0, (contents.width() - icon_size.width()) // 2),
                start_y,
                icon_size.width(),
                icon_size.height(),
            )
            text_rect = QtCore.QRect(
                contents.x(),
                icon_rect.bottom() + spacing,
                contents.width(),
                text_height,
            )

        icon_mode = QtGui.QIcon.Normal
        if not (option.state & QtWidgets.QStyle.State_Enabled):
            icon_mode = QtGui.QIcon.Disabled
        elif option.state & QtWidgets.QStyle.State_MouseOver:
            icon_mode = QtGui.QIcon.Active

        icon_state = QtGui.QIcon.On if option.state & QtWidgets.QStyle.State_On else QtGui.QIcon.Off

        if not icon.isNull() and icon_rect.isValid():
            icon.paint(painter, icon_rect, QtCore.Qt.AlignCenter, icon_mode, icon_state)

        if text and text_rect.isValid():
            alignment = QtCore.Qt.AlignCenter
            if style == QtCore.Qt.ToolButtonTextBesideIcon and icon_rect.isValid():
                alignment = QtCore.Qt.AlignVCenter | QtCore.Qt.AlignLeft
            elif style == QtCore.Qt.ToolButtonTextUnderIcon:
                alignment = QtCore.Qt.AlignHCenter | QtCore.Qt.AlignTop
            painter.setPen(self.palette().color(QtGui.QPalette.ButtonText))
            painter.drawText(text_rect, alignment, text)
