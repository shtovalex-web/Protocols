# -*- coding: utf-8 -*-
"""
Локальная синхронизация Linux-копии после правок в основном проекте.

  python tools/sync_linux_local.py
  python tools/sync_linux_local.py --prepare-only
  python tools/sync_linux_local.py --if-staged   # pre-commit / git hook

Обновляет:
  - linux_port/app/          (prepare.py)
  - ProtocolOOT_linux_build/ (pack_linux_build.py, по умолчанию)
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
LINUX_PORT = ROOT / "linux_port"
PREPARE = LINUX_PORT / "prepare.py"
PACK = ROOT / "tools" / "pack_linux_build.py"

if str(LINUX_PORT) not in sys.path:
    sys.path.insert(0, str(LINUX_PORT))

from sync_util import path_affects_linux_app  # noqa: E402


def _git_staged_paths() -> list[str]:
    r = subprocess.run(
        ["git", "diff", "--cached", "--name-only", "--diff-filter=ACMR"],
        cwd=str(ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if r.returncode != 0:
        return []
    return [ln.strip() for ln in (r.stdout or "").splitlines() if ln.strip()]


def _git_unstaged_and_staged_paths() -> list[str]:
    paths: list[str] = []
    for args in (
        ["git", "diff", "--name-only", "--diff-filter=ACMR"],
        ["git", "diff", "--cached", "--name-only", "--diff-filter=ACMR"],
        ["git", "ls-files", "--others", "--exclude-standard"],
    ):
        r = subprocess.run(
            args,
            cwd=str(ROOT),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if r.returncode == 0:
            paths.extend(ln.strip() for ln in (r.stdout or "").splitlines() if ln.strip())
    return paths


def needs_linux_sync(paths: list[str]) -> bool:
    return any(path_affects_linux_app(p) for p in paths)


def run_prepare(*, quiet: bool = False) -> int:
    if not PREPARE.is_file():
        print(f"Ошибка: нет {PREPARE}", file=sys.stderr)
        return 1
    cmd = [sys.executable, str(PREPARE)]
    if quiet:
        r = subprocess.run(cmd, cwd=str(ROOT), capture_output=True, text=True)
        if r.returncode != 0:
            if r.stdout:
                print(r.stdout, end="")
            if r.stderr:
                print(r.stderr, end="", file=sys.stderr)
        return r.returncode
    return subprocess.run(cmd, cwd=str(ROOT)).returncode


def run_pack(*, quiet: bool = False) -> int:
    if not PACK.is_file():
        print(f"Ошибка: нет {PACK}", file=sys.stderr)
        return 1
    cmd = [sys.executable, str(PACK), "--skip-prepare"]
    if quiet:
        r = subprocess.run(cmd, cwd=str(ROOT), capture_output=True, text=True)
        if r.returncode != 0:
            if r.stdout:
                print(r.stdout, end="")
            if r.stderr:
                print(r.stderr, end="", file=sys.stderr)
        return r.returncode
    return subprocess.run(cmd, cwd=str(ROOT)).returncode


def sync_linux_local(
    *,
    pack: bool = True,
    quiet: bool = False,
) -> int:
    if run_prepare(quiet=quiet) != 0:
        return 1
    if pack and run_pack(quiet=quiet) != 0:
        return 1
    if not quiet:
        print("Linux: linux_port/app/ обновлён" + ("; ProtocolOOT_linux_build/ упакован" if pack else ""))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Синхронизация linux_port/app и комплекта сборки")
    parser.add_argument(
        "--prepare-only",
        action="store_true",
        help="Только prepare.py, без ProtocolOOT_linux_build/",
    )
    parser.add_argument(
        "--if-staged",
        action="store_true",
        help="Выполнить только если в индексе git есть файлы, влияющие на Linux-копию",
    )
    parser.add_argument(
        "--if-changed",
        action="store_true",
        help="Выполнить только если есть изменённые/новые файлы источников (рабочая копия)",
    )
    parser.add_argument("--quiet", action="store_true", help="Минимум вывода (кроме ошибок)")
    args = parser.parse_args()

    if args.if_staged:
        paths = _git_staged_paths()
        if not needs_linux_sync(paths):
            if not args.quiet:
                print("Linux sync: пропуск (нет затронутых файлов в индексе)")
            return 0
    elif args.if_changed:
        paths = _git_unstaged_and_staged_paths()
        if not needs_linux_sync(paths):
            if not args.quiet:
                print("Linux sync: пропуск (нет затронутых файлов)")
            return 0

    return sync_linux_local(pack=not args.prepare_only, quiet=args.quiet)


if __name__ == "__main__":
    raise SystemExit(main())
