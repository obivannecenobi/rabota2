"""Entrypoint script for the application.

This module creates a virtual environment, installs dependencies and runs the
GUI application.  It performs a number of sanity checks and logs any failures
instead of abruptly terminating the program.
"""

from __future__ import annotations

import logging
import os
import subprocess
import sys
from pathlib import Path


LOGGER = logging.getLogger(__name__)
ROOT = Path(__file__).resolve().parent
VENV_DIR = ROOT / ".venv"
PYTHON = VENV_DIR / ("Scripts" if os.name == "nt" else "bin") / "python"
MIN_VERSION = (3, 10)


def _check_python() -> bool:
    """Validate python version."""

    if sys.version_info < MIN_VERSION:
        LOGGER.error(
            "Python %s.%s or newer is required", MIN_VERSION[0], MIN_VERSION[1]
        )
        return False
    return True


def main() -> int:
    """Run the application and return its exit code."""

    logging.basicConfig(level=logging.INFO)

    if not _check_python():
        return 1

    if not VENV_DIR.exists():
        try:
            subprocess.check_call([sys.executable, "-m", "venv", str(VENV_DIR)])
        except subprocess.CalledProcessError as exc:
            LOGGER.error("Failed to create virtualenv", exc_info=True)
            return exc.returncode

    try:
        subprocess.check_call(
            [
                str(PYTHON),
                "-m",
                "pip",
                "install",
                "--upgrade",
                "-r",
                str(ROOT / "requirements.txt"),
            ]
        )
    except subprocess.CalledProcessError as exc:
        LOGGER.error("Dependency installation failed", exc_info=True)
        return exc.returncode

    try:
        subprocess.check_call(
            [str(PYTHON), "-c", "import PySide6, blurwindow"]
        )
    except subprocess.CalledProcessError:
        LOGGER.error("Required dependencies PySide6 and/or blurwindow are missing")
        return 1

    os.environ.setdefault(
        "QT_QPA_PLATFORM",
        os.environ.get("QT_QPA_PLATFORM")
        or ("windows" if os.name == "nt" else "xcb"),
    )

    try:
        subprocess.check_call([str(PYTHON), str(ROOT / "app" / "main_window.py")])
    except subprocess.CalledProcessError as exc:
        LOGGER.error(
            "Application failed with exit code %s", exc.returncode, exc_info=True
        )
        return exc.returncode

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
