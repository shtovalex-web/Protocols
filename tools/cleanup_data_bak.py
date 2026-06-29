# -*- coding: utf-8
"""Удалить *.bak из data/ перед обновлением (старая схема отката).

Запуск:
    py -3 tools/cleanup_data_bak.py "D:\\путь\\к\\папке\\с\\exe"
    py -3 tools/cleanup_data_bak.py "D:\\путь\\к\\папке\\с\\exe\\data"
"""

from __future__ import annotations

import argparse
import os
import stat
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "ProtocolOHT_next"))

from update_bundle_files import DATA_SUBDIR_NAME  # noqa: E402


def resolve_data_dir(path: Path) -> Path:
    resolved = path.expanduser().resolve()
    if resolved.name.lower() == DATA_SUBDIR_NAME:
        return resolved
    candidate = resolved / DATA_SUBDIR_NAME
    if candidate.is_dir():
        return candidate
    return resolved


def cleanup_bak_files(data_dir: Path) -> list[str]:
    removed: list[str] = []
    if not data_dir.is_dir():
        return removed
    for path in sorted(data_dir.glob("*.bak")):
        if not path.is_file():
            continue
        try:
            os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
            path.unlink()
            removed.append(path.name)
        except OSError as error:
            print(f"FAIL: {path} — {error}", file=sys.stderr)
    return removed


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Удалить *.bak из data/ перед обновлением")
    parser.add_argument(
        "folder",
        type=Path,
        help="Папка с ProtocolOOT.exe или с подпапкой data/",
    )
    args = parser.parse_args(argv)
    data_dir = resolve_data_dir(args.folder)
    if not data_dir.is_dir():
        print(f"FAIL: нет каталога data: {data_dir}", file=sys.stderr)
        return 1
    removed = cleanup_bak_files(data_dir)
    if removed:
        print(f"Удалено в {data_dir}:")
        for name in removed:
            print(f"  - {name}")
    else:
        print(f"Файлов *.bak нет: {data_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
