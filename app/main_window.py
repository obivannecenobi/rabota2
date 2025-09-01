"""Main application window and entry point."""

from __future__ import annotations

import logging
import sys

from PySide6 import QtWidgets

from .utils import CONFIG, resolve_font_config
from .dialogs import (
    AnalyticsDialog,
    ReleaseDialog,
    SettingsDialog,
    StatsDialog,
    TopDialog,
)
from .widgets import StyledPushButton


LOGGER = logging.getLogger(__name__)


class MainWindow(QtWidgets.QMainWindow):
    """Main GUI window."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Web Novel Planner")

        central = QtWidgets.QWidget(self)
        layout = QtWidgets.QVBoxLayout(central)
        layout.addWidget(StyledPushButton("Hello"))
        self.setCentralWidget(central)


def main() -> int:
    """Start the Qt application."""

    logging.basicConfig(level=logging.INFO)
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication(sys.argv)
    try:
        resolve_font_config()
        window = MainWindow()
        window.show()
        return app.exec()
    except Exception:  # pragma: no cover - unexpected crash
        LOGGER.exception("Unhandled exception in GUI")
        print("Application error, see log for details", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

