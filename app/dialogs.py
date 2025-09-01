"""Reusable dialog classes."""

from __future__ import annotations

from PySide6 import QtWidgets


class ReleaseDialog(QtWidgets.QDialog):
    """Dialog for managing release schedule."""

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Release")


class StatsDialog(QtWidgets.QDialog):
    """Dialog for entering monthly statistics."""

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Statistics")


class AnalyticsDialog(QtWidgets.QDialog):
    """Yearly analytics dialog."""

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Analytics")


class TopDialog(QtWidgets.QDialog):
    """Dialog showing top metrics."""

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Top")


class SettingsDialog(QtWidgets.QDialog):
    """Application settings dialog."""

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Settings")

