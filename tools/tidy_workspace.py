# -*- coding: utf-8 -*-
"""
Уборка рабочей папки проекта: кэши, сборки, сгенерированные копии (без исходников и данных пользователя).

  python tools/tidy_workspace.py           — показать, что будет удалено
  python tools/tidy_workspace.py --apply   — удалить
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# Каталоги целиком (относительно корня проекта).
REMOVE_DIRS = (
    "__pycache__",
    ".pytest_cache",
    ".ruff_cache",
    "_pyinstaller_build",
    "_pyinstaller_build_onefile",
    "ProtocolOHT_onefile",
    "ProtocolOHT_app",
    "linux_port/app",
    "linux_port/.venv",
    "linux_port/ProtocolOHT_linux_dist",
    "linux_port/_pyinstaller_build_linux",
)

# Шаблоны имён каталогов/файлов в корне (ib_*, распаковки ИБ).
REMOVE_GLOBS = (
    "ib_*",
    "ProtocolOOT.spec",
    "linux_port/ProtocolOOT.spec",
)


def _iter_remove_targets() -> list[Path]:
    found: list[Path] = []
    seen: set[Path] = set()

    def add(p: Path) -> None:
        rp = p.resolve()
        if rp in seen or not p.exists():
            return
        seen.add(rp)
        found.append(p)

    for name in REMOVE_DIRS:
        add(ROOT / name)

    for pattern in REMOVE_GLOBS:
        for p in ROOT.glob(pattern):
            add(p)

    for p in ROOT.rglob("__pycache__"):
        add(p)

    return sorted(found, key=lambda x: len(x.parts), reverse=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="Уборка кэшей и артефактов сборки")
    parser.add_argument(
        "--apply",
        action="store_true",
        help="Удалить найденное (без флага — только список)",
    )
    args = parser.parse_args()
    targets = _iter_remove_targets()
    if not targets:
        print("Удалять нечего — рабочая папка уже чистая.")
        return 0

    print("Найдено для удаления:")
    for p in targets:
        rel = p.relative_to(ROOT) if p.is_relative_to(ROOT) else p
        kind = "каталог" if p.is_dir() else "файл"
        print(f"  [{kind}] {rel}")

    if not args.apply:
        print()
        print("Для удаления: python tools/tidy_workspace.py --apply")
        return 0

    errors = 0
    for p in targets:
        try:
            if p.is_dir():
                shutil.rmtree(p)
            else:
                p.unlink()
        except OSError as e:
            print(f"Ошибка: {p}: {e}", file=sys.stderr)
            errors += 1
    print(f"Готово. Удалено объектов: {len(targets) - errors}")
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
