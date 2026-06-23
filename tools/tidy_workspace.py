# -*- coding: utf-8 -*-
"""
Уборка рабочей папки проекта: кэши, сборки, сгенерированные копии (без исходников и данных пользователя).

  python tools/tidy_workspace.py           — показать, что будет удалено
  python tools/tidy_workspace.py --apply   — удалить

Не трогает: protocols.db, Data_base.xlsx, Programs_base.xlsx, Protokol/, Mintrud/ (рабочие данные).
"""

from __future__ import annotations

import argparse
import os
import shutil
import stat
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ETALON = ROOT / "эталон_сборки"
BUNDLE = ROOT / "bundle"

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
    "linux_port/.venv-linux",
    "ProtocolOOT_linux_build",
    "ProtocolOOT_linux_build_test_tmp",
    "linux_port/release/out_linux",
    "linux_port/release/_build",
    "linux_port/ProtocolOHT_linux_dist",
    "linux_port/_pyinstaller_build_linux",
)

# Шаблоны имён каталогов/файлов в корне (ib_*, распаковки ИБ).
REMOVE_GLOBS = (
    "ib_*",
    "ProtocolOOT.spec",
    "linux_port/ProtocolOOT.spec",
)

# Дубликаты комплекта в корне репозитория (канон — bundle/).
ROOT_BUNDLE_DUPLICATES = (
    "default_protocol.docx",
    "default_protocol_tehnicheskiy.docx",
    "educated_person_import_v1.0.9.xsd",
    "icon.ico",
)

# Случайно появившиеся в корне эталона (комплект — в data/).
ETALON_RUNTIME_JUNK = (
    "protocols.db",
    "last_protocol_no.json",
    "protocol_errors_journal.txt",
    "desktop.ini",
)

ETALON_DUPLICATE_NAMES = (
    "default_protocol.docx",
    "default_protocol_tehnicheskiy.docx",
    "FAQ.md",
    "icon.ico",
    "Шаблон_Минтруд_XSD_УМН.xlsx",
    "ИНСТРУКЦИЯ_для_пользователя.md",
    "ИНСТРУКЦИЯ_оформление_протоколов_Минтруд.docx",
    "ИНСТРУКЦИЯ_оформление_протоколов_Минтруд.md",
    "ПОДРОБНАЯ_ИНСТРУКЦИЯ_для_пользователя.docx",
    "ПОДРОБНАЯ_ИНСТРУКЦИЯ_для_пользователя.md",
)

_SKIP_RGLOB_DIR_NAMES = frozenset({"venv", ".venv"})


def _rmtree_robust(target: Path) -> None:
    """Удаление каталога на Windows (read-only, блокировка файла)."""

    def _onexc(func, path, exc):
        if isinstance(exc, PermissionError):
            try:
                os.chmod(path, stat.S_IWRITE)
                func(path)
            except OSError as e2:
                raise exc from e2
        else:
            raise exc

    if not target.exists():
        return
    last: OSError | None = None
    for _ in range(5):
        try:
            if sys.version_info >= (3, 12):
                shutil.rmtree(target, onexc=_onexc)
            else:

                def _legacy(func, path, exc_info):
                    _onexc(func, path, exc_info[1])

                shutil.rmtree(target, onerror=_legacy)
            return
        except OSError as e:
            last = e
            time.sleep(0.25)
    if last:
        raise last


def _should_skip_rglob(p: Path) -> bool:
    return any(part in _SKIP_RGLOB_DIR_NAMES for part in p.parts)


def _iter_root_bundle_duplicates() -> list[Path]:
    found: list[Path] = []
    for name in ROOT_BUNDLE_DUPLICATES:
        root_file = ROOT / name
        bundle_file = BUNDLE / name
        if root_file.is_file() and bundle_file.is_file():
            found.append(root_file)
    return found


def iter_etalon_stale_paths(etalon: Path = ETALON) -> list[Path]:
    """Файлы в корне эталона, которые не должны там оставаться после запуска программы."""
    if not etalon.is_dir():
        return []
    data = etalon / "data"
    found: list[Path] = []
    for name in ETALON_RUNTIME_JUNK:
        p = etalon / name
        if p.is_file():
            found.append(p)
    for name in ETALON_DUPLICATE_NAMES:
        p = etalon / name
        if not p.is_file():
            continue
        if data.is_dir() and (data / name).is_file():
            found.append(p)
            continue
        if name == "FAQ.md" and (data / "FAQ.txt").is_file():
            found.append(p)
    stale_cache = etalon / "__pycache__"
    if stale_cache.is_dir():
        found.append(stale_cache)
    return found


def cleanup_etalon_root(*, apply: bool, etalon: Path = ETALON) -> list[Path]:
    """Убрать мусор и дубликаты data/ из корня эталон_сборки/."""
    targets = iter_etalon_stale_paths(etalon)
    if apply:
        for p in targets:
            try:
                if p.is_dir():
                    _rmtree_robust(p)
                else:
                    p.unlink()
            except OSError as e:
                print(f"Ошибка эталона: {p}: {e}", file=sys.stderr)
    return targets


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
        if _should_skip_rglob(p):
            continue
        add(p)

    for p in _iter_root_bundle_duplicates():
        add(p)

    for p in iter_etalon_stale_paths():
        add(p)

    return sorted(found, key=lambda x: len(x.parts), reverse=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="Уборка кэшей и артефактов сборки")
    parser.add_argument(
        "--apply",
        action="store_true",
        help="Удалить найденное (без флага — только список)",
    )
    parser.add_argument(
        "--etalon-only",
        action="store_true",
        help="Только очистка корня эталон_сборки/",
    )
    args = parser.parse_args()

    if args.etalon_only:
        targets = iter_etalon_stale_paths()
        if not targets:
            print("Эталон: убирать нечего.")
            return 0
        print("Эталон — найдено для удаления:")
        for p in targets:
            rel = p.relative_to(ROOT) if p.is_relative_to(ROOT) else p
            kind = "каталог" if p.is_dir() else "файл"
            print(f"  [{kind}] {rel}")
        if not args.apply:
            print()
            print("Для удаления: python tools/tidy_workspace.py --etalon-only --apply")
            return 0
        cleanup_etalon_root(apply=True)
        print(f"Эталон: удалено объектов: {len(targets)}")
        return 0

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
                _rmtree_robust(p)
            else:
                if not os.access(p, os.W_OK):
                    try:
                        os.chmod(p, stat.S_IWRITE)
                    except OSError:
                        pass
                p.unlink()
        except OSError as e:
            print(f"Ошибка: {p}: {e}", file=sys.stderr)
            errors += 1
    print(f"Готово. Удалено объектов: {len(targets) - errors}")
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
