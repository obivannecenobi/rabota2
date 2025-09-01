from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

if sys.version_info < (3, 13):
    sys.exit('Требуется Python 3.13+')

ROOT = Path(__file__).resolve().parent
VENV_DIR = ROOT / '.venv'
PYTHON = VENV_DIR / ('Scripts' if os.name == 'nt' else 'bin') / 'python'


def main() -> int:
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
        print(f'Не удалось установить зависимости: {exc}', file=sys.stderr)
        return exc.returncode

    result = subprocess.run(
        [str(PYTHON), '-X', 'faulthandler', str(ROOT / 'app' / 'main.py')],
        stderr=subprocess.PIPE,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        if result.stderr:
            print(result.stderr, file=sys.stderr, end="")
        print(
            f"Application failed with exit code {result.returncode}",
            file=sys.stderr,
        )
        return result.returncode

    return 0


if __name__ == '__main__':
    raise SystemExit(main())
