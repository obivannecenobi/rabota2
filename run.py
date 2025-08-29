from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
VENV_DIR = ROOT / '.venv'
PYTHON = VENV_DIR / ('Scripts' if os.name == 'nt' else 'bin') / 'python'


def main() -> None:
    if not VENV_DIR.exists():
        subprocess.check_call([sys.executable, '-m', 'venv', str(VENV_DIR)])

    try:
        subprocess.check_call(
            [
                str(PYTHON),
                '-m',
                'pip',
                'install',
                '--upgrade',
                '-r',
                str(ROOT / 'requirements.txt'),
            ]
        )
    except subprocess.CalledProcessError as exc:
        print(f"Failed to install requirements: {exc}", file=sys.stderr)
        sys.exit(exc.returncode)

    subprocess.check_call([str(PYTHON), str(ROOT / 'app' / 'main.py')])


if __name__ == '__main__':
    main()
