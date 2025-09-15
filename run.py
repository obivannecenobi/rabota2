from __future__ import annotations

import argparse
import importlib.util
import os
import subprocess
import sys
from pathlib import Path

if sys.version_info < (3, 11):
    sys.exit('Требуется Python 3.11+')

ROOT = Path(__file__).resolve().parent
VENV_DIR = ROOT / '.venv'
PYTHON = VENV_DIR / ('Scripts' if os.name == 'nt' else 'bin') / 'python'
LOG_DIR = ROOT / 'logs'
LOG_PATH = LOG_DIR / 'startup.log'


def ensure_gui_dependencies() -> None:
    """Проверить наличие tkinter и customtkinter и подсказать установку."""

    has_tkinter = importlib.util.find_spec("tkinter") is not None
    has_customtkinter = importlib.util.find_spec("customtkinter") is not None

    if not has_tkinter:
        message = (
            "tkinter не найден. Установите поддержку Tcl/Tk для вашей ОС."
        )
        if sys.platform.startswith("win"):
            message += (
                " В установщике Python выберите «Modify > tcl/tk and IDLE»."
            )
        elif sys.platform == "darwin":
            message += (
                " На macOS установите Python с сайта python.org или добавьте"
                " пакет Tcl/Tk через Homebrew: `brew install tcl-tk`."
            )
        else:
            message += (
                " На Linux обычно достаточно установить пакет python3-tk"
                " через менеджер пакетов, например: `sudo apt install python3-tk`."
            )
        print(message)

    if not has_customtkinter:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "customtkinter"]
        )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--install",
        action="store_true",
        help="переустановить зависимости в виртуальном окружении",
    )
    args = parser.parse_args()

    created = False
    if not VENV_DIR.exists():
        subprocess.check_call([sys.executable, "-m", "venv", str(VENV_DIR)])
        created = True

    if created or args.install:
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
            print(f"Не удалось установить зависимости: {exc}", file=sys.stderr)
            return exc.returncode

    ensure_gui_dependencies()

    LOG_DIR.mkdir(parents=True, exist_ok=True)
    with LOG_PATH.open("w", encoding="utf-8") as log_file:
        proc = subprocess.Popen(
            [str(PYTHON), str(ROOT / "app" / "main.py")],
            stdout=log_file,
            stderr=log_file,
            text=True,
        )
        proc.wait()

    if proc.returncode != 0:
        print(
            f"Application failed with exit code {proc.returncode}. See {LOG_PATH} for details",
            file=sys.stderr,
        )
        return proc.returncode

    return 0


if __name__ == '__main__':
    raise SystemExit(main())
