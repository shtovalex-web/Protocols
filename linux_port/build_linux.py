# -*- coding: utf-8 -*-
"""Обёртка: python linux_port/build_linux.py → release/build_release_linux.py --local."""

from __future__ import annotations

import runpy
import sys
from pathlib import Path

LINUX_PORT = Path(__file__).resolve().parent
TARGET = LINUX_PORT / "release" / "build_release_linux.py"


def _fail(message: str) -> None:
    print(message, file=sys.stderr)
    raise SystemExit(1)


if __name__ == "__main__":
    if not TARGET.is_file():
        _fail(
            f"Не найден {TARGET.relative_to(LINUX_PORT.parent)}.\n"
            "Обновите репозиторий (git pull) или выполните сборку из актуальной ветки linux."
        )
    if not (LINUX_PORT / "app" / "main.py").is_file():
        _fail(
            "Нет linux_port/app/main.py.\n"
            "Выполните: python3 linux_port/prepare.py\n"
            "На ветке linux: git clone -b linux …"
        )

    extras = list(sys.argv[1:])
    if not any(x in extras for x in ("--local", "-o", "--output", "--binary-only")):
        extras.insert(0, "--local")
    sys.argv = [str(TARGET), *extras]
    runpy.run_path(str(TARGET), run_name="__main__")
