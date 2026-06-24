# -*- coding: utf-8 -*-
"""Публикация manifest.json и копии .exe на сетевую шару."""

from __future__ import annotations

import argparse
import json
import shutil
import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
_NEXT = ROOT / "ProtocolOHT_next"
sys.path.insert(0, str(_NEXT))

from update_manifest import sha256_file  # noqa: E402


def _default_exe_name() -> str:
    return "ProtocolOOT.exe"


def publish(
    *,
    exe_path: Path,
    version: str,
    share_root: Path,
    changes: list[str],
    mandatory: bool,
    released: str,
) -> Path:
    if not exe_path.is_file():
        msg = f"EXE not found: {exe_path}"
        raise SystemExit(msg)

    target_dir = share_root / "windows" / version
    target_dir.mkdir(parents=True, exist_ok=True)
    target_exe = target_dir / _default_exe_name()
    shutil.copy2(exe_path, target_exe)

    digest = sha256_file(target_exe)
    size = target_exe.stat().st_size
    manifest = {
        "latest_version": version,
        "released": released,
        "mandatory": mandatory,
        "windows": {
            "relative_path": f"windows/{version}/{target_exe.name}".replace("\\", "/"),
            "sha256": digest,
            "size": size,
        },
        "changes_short": changes,
    }
    manifest_path = share_root / "manifest.json"
    manifest_path.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(f"Payload: {target_exe} ({size} bytes, sha256={digest[:16]}...)")
    return manifest_path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Опубликовать обновление ProtocolOOT на шару")
    parser.add_argument("--exe", type=Path, required=True, help="Путь к собранному ProtocolOOT.exe")
    parser.add_argument("--version", required=True, help="Версия релиза, напр. 1.6.0")
    parser.add_argument(
        "--share-root",
        type=Path,
        required=True,
        help=r"Корень шары, напр. \\SERVER\SOFT\ProtocolOOT",
    )
    parser.add_argument(
        "--change",
        action="append",
        default=[],
        dest="changes",
        help="Краткое описание изменения (можно несколько раз)",
    )
    parser.add_argument("--mandatory", action="store_true", help="Обязательное обновление")
    parser.add_argument(
        "--released",
        default=date.today().isoformat(),
        help="Дата релиза YYYY-MM-DD",
    )
    args = parser.parse_args(argv)

    manifest_path = publish(
        exe_path=args.exe.expanduser().resolve(),
        version=args.version.strip(),
        share_root=args.share_root.expanduser().resolve(),
        changes=[str(c) for c in args.changes if str(c).strip()],
        mandatory=bool(args.mandatory),
        released=str(args.released).strip(),
    )
    print(f"Manifest: {manifest_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
