# -*- coding: utf-8 -*-
"""
Проверка после правок и запуск приложения.

  python tools/verify_project.py              — ruff + импорты + запуск main.py (новое окно)
  python tools/verify_project.py --no-launch — только ruff и импорты (CI, быстрая проверка)
  python tools/verify_project.py --no-linux-sync — без обновления linux_port/app

Устаревший алиас: --gui (то же, что поведение по умолчанию).

Рекомендуется: pip install ruff (см. requirements-build.txt).
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
_NEXT = ROOT / "ProtocolOHT_next"
_LINUX_PORT = ROOT / "linux_port"


def run_ruff() -> None:
    r = subprocess.run(
        [sys.executable, "-m", "ruff", "check", "."],
        cwd=ROOT,
        capture_output=True,
        text=True,
    )
    if r.returncode == 0:
        print("ruff check: OK")
        return
    err = f"{r.stderr or ''}{r.stdout or ''}"
    if "No module named" in err and "ruff" in err:
        print("ruff не установлен — пропуск (pip install ruff)")
        return
    if r.stdout:
        print(r.stdout, end="")
    if r.stderr:
        print(r.stderr, end="", file=sys.stderr)
    raise SystemExit(r.returncode)


def smoke_imports() -> None:
    """Те же пути, что в main.py — без создания окна tkinter."""
    for _p in (str(ROOT), str(_NEXT)):
        while _p in sys.path:
            sys.path.remove(_p)
    sys.path.insert(0, str(_NEXT))
    sys.path.insert(1, str(ROOT))

    import commission_admin  # noqa: F401
    import employees_io  # noqa: F401
    import protocol_db  # noqa: F401
    from protocol_docx import build_filled_protocol_document  # noqa: F401
    from protocol_ui import ProtocolApp  # noqa: F401

    print("Импорты (commission_admin, employees_io, protocol_db, protocol_docx, protocol_ui): OK")


def launch_app() -> None:
    subprocess.Popen(
        [sys.executable, str(ROOT / "main.py")],
        cwd=ROOT,
        close_fds=True,
    )
    print("Запущен отдельный процесс main.py — проверьте окно приложения.")


def run_unit_tests() -> None:
    r = subprocess.run(
        [sys.executable, "-m", "unittest", "discover", "-s", "tests", "-v"],
        cwd=ROOT,
    )
    if r.returncode != 0:
        raise SystemExit(r.returncode)
    print("unittest (tests/): OK")


def sync_linux_after_verify(*, pack: bool = True) -> None:
    """Обновить linux_port/app/ и комплект ProtocolOOT_linux_build/ при изменённых источниках."""
    sync_dir = ROOT / "tools"
    if str(sync_dir) not in sys.path:
        sys.path.insert(0, str(sync_dir))
    if str(_LINUX_PORT) not in sys.path:
        sys.path.insert(0, str(_LINUX_PORT))

    from sync_linux_local import (  # noqa: PLC0415
        _git_unstaged_and_staged_paths,
        needs_linux_sync,
        sync_linux_local,
    )

    paths = _git_unstaged_and_staged_paths()
    if not needs_linux_sync(paths):
        print("Linux sync: пропуск (источники приложения не менялись)")
        return
    if sync_linux_local(pack=pack, quiet=False) != 0:
        raise SystemExit(1)


def main() -> int:
    argv = list(sys.argv[1:])
    no_launch = False
    no_linux_sync = False
    linux_prepare_only = False
    for flag in ("--no-launch", "-n"):
        while flag in argv:
            argv.remove(flag)
            no_launch = True
    for flag in ("--no-linux-sync",):
        while flag in argv:
            argv.remove(flag)
            no_linux_sync = True
    for flag in ("--linux-prepare-only",):
        while flag in argv:
            argv.remove(flag)
            linux_prepare_only = True
    for legacy in ("--gui",):
        while legacy in argv:
            argv.remove(legacy)
    if argv:
        print("Неизвестные аргументы:", argv, file=sys.stderr)
        return 2

    run_ruff()
    smoke_imports()
    run_unit_tests()
    if not no_linux_sync:
        sync_linux_after_verify(pack=not linux_prepare_only)
    if not no_launch:
        launch_app()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
