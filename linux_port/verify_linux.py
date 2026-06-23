# -*- coding: utf-8 -*-
"""
Проверка Linux-копии (linux_port/app/).

  python linux_port/verify_linux.py              — ruff + импорты + запуск main.py
  python linux_port/verify_linux.py --no-launch — только ruff и импорты
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

LINUX_PORT = Path(__file__).resolve().parent
APP = LINUX_PORT / "app"
_NEXT = APP / "ProtocolOHT_next"


def run_ruff() -> None:
    from ruff_linux import run_ruff_on_app

    r = run_ruff_on_app(linux_port=LINUX_PORT, app_dir=APP, cwd=LINUX_PORT, capture=True)
    if r.returncode == 0:
        print("ruff check (app/): OK")
        return
    err = f"{r.stderr or ''}{r.stdout or ''}"
    if "No module named" in err and "ruff" in err:
        print("ruff не установлен — пропуск")
        return
    if r.stdout:
        print(r.stdout, end="")
    if r.stderr:
        print(r.stderr, end="", file=sys.stderr)
    raise SystemExit(r.returncode)


def smoke_imports() -> None:
    if not APP.is_dir():
        raise SystemExit(f"Нет каталога {APP} — выполните: python linux_port/prepare.py")

    for _p in (str(APP), str(_NEXT)):
        while _p in sys.path:
            sys.path.remove(_p)
    sys.path.insert(0, str(_NEXT))
    sys.path.insert(1, str(APP))

    import bundle_integration  # noqa: F401
    import commission_admin  # noqa: F401
    import employees_io  # noqa: F401
    import protocol_db  # noqa: F401
    from protocol_docx import build_filled_protocol_document  # noqa: F401
    from protocol_output import cyrillic_ttf_candidates, _libreoffice_executable  # noqa: F401
    from protocol_ui import ProtocolApp  # noqa: F401

    print("Импорты Linux-копии: OK")
    print(f"  LibreOffice в PATH: {_libreoffice_executable() or 'не найден (нормально на Windows)'}")


def launch_app() -> None:
    subprocess.run([sys.executable, str(APP / "main.py")], cwd=str(APP), check=False)


def main() -> int:
    no_launch = "--no-launch" in sys.argv
    run_ruff()
    smoke_imports()
    if no_launch:
        print("verify_linux: OK (--no-launch)")
        return 0
    print("Запуск main.py...")
    launch_app()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
