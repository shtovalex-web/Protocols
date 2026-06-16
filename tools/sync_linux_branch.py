# -*- coding: utf-8 -*-
"""
Синхронизация ветки linux с main (Windows): merge + полная копия в linux_port/app/.

Из корня репозитория:
    python tools/sync_linux_branch.py
    python tools/sync_linux_branch.py --push
    python tools/sync_linux_branch.py --source main --message "sync после правок Минтруда"

На ветке linux в git хранится готовая копия linux_port/app/ (на main она в .gitignore).
После push в main GitHub Actions (sync-linux.yml) обновляет ветку linux автоматически.
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
LINUX_BRANCH = "linux"
APP_DIR = ROOT / "linux_port" / "app"
MANIFEST = APP_DIR / ".linux_sync_manifest.json"


def _run(cmd: list[str], *, check: bool = True, capture: bool = False) -> subprocess.CompletedProcess[str]:
    print("+", " ".join(cmd))
    return subprocess.run(
        cmd,
        cwd=str(ROOT),
        check=check,
        capture_output=capture,
        text=True,
        encoding="utf-8",
        errors="replace",
    )


def _git_out(*args: str) -> str:
    r = _run(["git", *args], capture=True)
    return (r.stdout or "").strip()


def _current_branch() -> str:
    return _git_out("branch", "--show-current")


def _branch_exists(name: str) -> bool:
    r = subprocess.run(
        ["git", "rev-parse", "--verify", name],
        cwd=str(ROOT),
        capture_output=True,
    )
    return r.returncode == 0


def _working_tree_clean() -> bool:
    r1 = subprocess.run(["git", "diff", "--quiet"], cwd=str(ROOT))
    r2 = subprocess.run(["git", "diff", "--cached", "--quiet"], cwd=str(ROOT))
    return r1.returncode == 0 and r2.returncode == 0


def _run_prepare() -> int:
    return _run([sys.executable, str(ROOT / "linux_port" / "prepare.py")]).returncode


def _run_verify() -> int:
    verify = ROOT / "linux_port" / "verify_linux.py"
    if not verify.is_file():
        return 0
    return _run([sys.executable, str(verify), "--no-launch"]).returncode


def _write_manifest(source_branch: str) -> None:
    APP_DIR.mkdir(parents=True, exist_ok=True)
    try:
        source_commit = _git_out("rev-parse", source_branch)
    except SystemExit:
        source_commit = _git_out("rev-parse", "HEAD")
    payload = {
        "synced_at_utc": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "source_branch": source_branch,
        "source_commit": source_commit,
        "linux_branch": LINUX_BRANCH,
    }
    MANIFEST.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _commit_linux_app(message: str) -> bool:
    import shutil as _shutil

    for cache in APP_DIR.rglob("__pycache__"):
        if cache.is_dir():
            _shutil.rmtree(cache, ignore_errors=True)
    _run(["git", "add", "-f", "linux_port/app/"])
    _run(["git", "add", "linux_port/"])
    r = subprocess.run(["git", "diff", "--cached", "--quiet"], cwd=str(ROOT))
    if r.returncode == 0:
        print("Нет изменений для коммита на ветке linux.")
        return False
    _run(["git", "commit", "-m", message])
    return True


def main() -> int:
    parser = argparse.ArgumentParser(description="Синхронизация ветки linux с main")
    parser.add_argument(
        "--source",
        default="main",
        help="Ветка-источник (Windows), по умолчанию main",
    )
    parser.add_argument(
        "--message",
        default="",
        help="Сообщение коммита на linux (по умолчанию — sync: …)",
    )
    parser.add_argument("--push", action="store_true", help="git push origin linux")
    parser.add_argument(
        "--allow-dirty",
        action="store_true",
        help="Разрешить незакоммиченные правки на source (prepare читает с диска)",
    )
    parser.add_argument(
        "--no-merge",
        action="store_true",
        help="Не делать merge source → linux (только prepare и commit app/)",
    )
    args = parser.parse_args()

    if not (ROOT / "linux_port" / "prepare.py").is_file():
        print("Ошибка: нет linux_port/prepare.py", file=sys.stderr)
        return 1

    if not args.allow_dirty and not _working_tree_clean():
        print(
            "Рабочая копия с незакоммиченными изменениями.\n"
            "Закоммитьте, сделайте stash или укажите --allow-dirty "
            "(prepare возьмёт файлы с диска).",
            file=sys.stderr,
        )
        return 1

    original = _current_branch()
    source = args.source

    if _branch_exists("origin/main") or _branch_exists("remotes/origin/main"):
        try:
            _run(["git", "fetch", "origin"], check=False)
        except SystemExit:
            pass

    if not _branch_exists(source):
        print(f"Ошибка: нет ветки {source}", file=sys.stderr)
        return 1

    if _branch_exists(LINUX_BRANCH):
        _run(["git", "checkout", LINUX_BRANCH])
        if not args.no_merge and _branch_exists(source):
            merge_target = source
            if _branch_exists(f"origin/{source}"):
                merge_target = f"origin/{source}"
            r = subprocess.run(
                ["git", "merge", merge_target, "--no-edit", "-m", f"merge {source} into linux"],
                cwd=str(ROOT),
            )
            if r.returncode != 0:
                print("Конфликт merge. Разрешите вручную и закоммитьте.", file=sys.stderr)
                return 1
    else:
        print(f"Создание ветки {LINUX_BRANCH} от {source}")
        _run(["git", "checkout", "-b", LINUX_BRANCH, source])

    if _run_prepare() != 0:
        return 1
    if _run_verify() != 0:
        return 1

    _write_manifest(source)
    msg = args.message.strip() or f"sync: linux app from {source} ({datetime.now().strftime('%Y-%m-%d %H:%M')})"
    _commit_linux_app(msg)

    if args.push:
        _run(["git", "push", "-u", "origin", LINUX_BRANCH], check=False)

    if original and original != LINUX_BRANCH and _branch_exists(original):
        _run(["git", "checkout", original])

    print()
    print(f"Ветка {LINUX_BRANCH} обновлена.")
    print("На Linux: git clone -b linux … && cd linux_port && ./install_deps.sh && ./run.sh")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
