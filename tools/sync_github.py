# -*- coding: utf-8 -*-
"""Синхронизация проекта с GitHub: pull → add → commit → push.

Запуск из корня:
    py -3 tools/sync_github.py
    py -3 tools/sync_github.py -m "описание изменений"
    py -3 tools/sync_github.py --pull-only
    py -3 tools/sync_github.py --push-only

Или двойной щелчок: sync_github.bat
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def _run(cmd: list[str], *, check: bool = True) -> int:
    print("+", " ".join(cmd))
    r = subprocess.run(cmd, cwd=str(ROOT))
    if check and r.returncode != 0:
        raise SystemExit(r.returncode)
    return r.returncode


def _has_remote() -> bool:
    r = subprocess.run(
        ["git", "remote", "get-url", "origin"],
        cwd=str(ROOT),
        capture_output=True,
    )
    return r.returncode == 0


def _has_changes_to_commit() -> bool:
    r = subprocess.run(
        ["git", "diff", "--cached", "--quiet"],
        cwd=str(ROOT),
    )
    staged = r.returncode != 0
    r2 = subprocess.run(
        ["git", "diff", "--quiet"],
        cwd=str(ROOT),
    )
    unstaged = r2.returncode != 0
    r3 = subprocess.run(
        ["git", "ls-files", "--others", "--exclude-standard", "--directory"],
        cwd=str(ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    untracked = bool(r3.stdout.strip())
    return staged or unstaged or untracked


def main() -> int:
    parser = argparse.ArgumentParser(description="Синхронизация с GitHub (origin)")
    parser.add_argument(
        "-m",
        "--message",
        default="",
        help="Сообщение коммита (по умолчанию — дата/время)",
    )
    parser.add_argument("--pull-only", action="store_true", help="Только git pull --rebase")
    parser.add_argument("--push-only", action="store_true", help="Только git push")
    parser.add_argument(
        "--no-pull",
        action="store_true",
        help="Не выполнять pull перед коммитом",
    )
    args = parser.parse_args()

    if not _has_remote():
        print("Ошибка: не настроен remote origin.", file=sys.stderr)
        return 1

    branch = subprocess.run(
        ["git", "branch", "--show-current"],
        cwd=str(ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=True,
    ).stdout.strip() or "main"

    if not args.push_only and not args.no_pull:
        _run(["git", "pull", "--rebase", "origin", branch], check=False)

    if args.pull_only:
        return 0

    if args.push_only:
        return _run(["git", "push", "origin", branch], check=True)

    _run(["git", "add", "-A"])
    if not _has_changes_to_commit():
        print("Нет изменений для коммита — только push.")
        return _run(["git", "push", "origin", branch], check=True)

    msg = args.message.strip()
    if not msg:
        msg = f"Синхронизация {datetime.now().strftime('%Y-%m-%d %H:%M')}"

    _run(["git", "commit", "-m", msg])
    return _run(["git", "push", "origin", branch], check=True)


if __name__ == "__main__":
    raise SystemExit(main())
