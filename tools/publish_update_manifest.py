# -*- coding: utf-8
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

from update_bundle_files import DATA_REPLACE_FILENAMES, build_data_manifest_entries  # noqa: E402
from update_manifest import sha256_file  # noqa: E402
from windows_app_bundle import WINDOWS_APP_ZIP_NAME, resolve_windows_payload  # noqa: E402


def _default_exe_name() -> str:
    return "ProtocolOOT.exe"


def verify_manifest_payload(manifest_path: Path) -> list[str]:
    """Проверяет, что файлы на шаре совпадают с manifest.json; список ошибок."""
    errors: list[str] = []
    try:
        payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        return [f"manifest: {error}"]

    version_dir = manifest_path.parent
    windows = payload.get("windows") if isinstance(payload, dict) else None
    if isinstance(windows, dict):
        rel = str(windows.get("relative_path", "")).strip()
        expected_sha = str(windows.get("sha256", "")).strip()
        try:
            expected_size = int(windows.get("size", 0))
        except (TypeError, ValueError):
            expected_size = 0
        exe_path = version_dir / rel
        errors.extend(_verify_file_entry(exe_path, expected_sha, expected_size, label=rel or "windows"))

    data_files = payload.get("data_files", []) if isinstance(payload, dict) else []
    if isinstance(data_files, list):
        for item in data_files:
            if not isinstance(item, dict):
                continue
            rel = str(item.get("relative_path", "")).strip()
            expected_sha = str(item.get("sha256", "")).strip()
            try:
                expected_size = int(item.get("size", 0))
            except (TypeError, ValueError):
                expected_size = 0
            file_path = version_dir / rel.replace("/", "\\")
            errors.extend(
                _verify_file_entry(file_path, expected_sha, expected_size, label=rel or "data_files[]")
            )
    return errors


def _verify_file_entry(
    path: Path,
    expected_sha: str,
    expected_size: int,
    *,
    label: str,
) -> list[str]:
    if not path.is_file():
        return [f"{label}: файл не найден: {path}"]
    actual_size = path.stat().st_size
    if actual_size != expected_size:
        return [f"{label}: размер {actual_size} != {expected_size}"]
    actual_sha = sha256_file(path)
    if actual_sha.lower() != expected_sha.lower():
        return [f"{label}: sha256 не совпадает"]
    return []


def _write_manifest(
    *,
    target_dir: Path,
    target_payload: Path,
    target_data: Path,
    version: str,
    changes: list[str],
    mandatory: bool,
    released: str,
) -> Path:
    digest = sha256_file(target_payload)
    size = target_payload.stat().st_size
    data_entries = build_data_manifest_entries(
        data_src_dir=target_data,
        paths_relative_to_version_dir=True,
    )
    manifest = {
        "latest_version": version,
        "released": released,
        "mandatory": mandatory,
        "windows": {
            "relative_path": target_payload.name,
            "sha256": digest,
            "size": size,
        },
        "changes_short": changes,
        "data_files": data_entries,
    }
    manifest_path = target_dir / "manifest.json"
    manifest_path.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    errors = verify_manifest_payload(manifest_path)
    if errors:
        manifest_path.unlink(missing_ok=True)
        details = "\n".join(f"  - {item}" for item in errors)
        msg = f"Manifest verification failed:\n{details}"
        raise SystemExit(msg)
    return manifest_path


def refresh_manifest(
    *,
    version_dir: Path,
    version: str,
    changes: list[str],
    mandatory: bool,
    released: str,
) -> Path:
    """Пересобирает manifest.json по уже лежащим на шаре exe/zip и data/ (без копирования)."""
    target_dir = version_dir.expanduser().resolve()
    target_payload = resolve_windows_payload(target_dir)
    if not target_payload.is_file():
        msg = f"Windows payload not found in version dir: {target_payload}"
        raise SystemExit(msg)
    target_data = target_dir / "data"
    manifest_path = _write_manifest(
        target_dir=target_dir,
        target_payload=target_payload,
        target_data=target_data,
        version=version,
        changes=changes,
        mandatory=mandatory,
        released=released,
    )
    print(f"Refreshed manifest: {manifest_path}")
    return manifest_path


def publish(
    *,
    exe_path: Path,
    version: str,
    share_root: Path,
    changes: list[str],
    mandatory: bool,
    released: str,
    data_src_dir: Path | None,
    app_zip: Path | None = None,
) -> Path:
    if not exe_path.is_file():
        msg = f"EXE not found: {exe_path}"
        raise SystemExit(msg)

    target_dir = share_root / "windows" / version
    target_dir.mkdir(parents=True, exist_ok=True)
    target_exe = target_dir / _default_exe_name()
    shutil.copy2(exe_path, target_exe)

    if app_zip is not None and app_zip.is_file():
        target_payload = target_dir / WINDOWS_APP_ZIP_NAME
        shutil.copy2(app_zip, target_payload)
    else:
        target_payload = target_exe

    data_dir = (data_src_dir or exe_path.parent / "data").expanduser().resolve()
    target_data = target_dir / "data"
    copied_data = 0
    for name in DATA_REPLACE_FILENAMES:
        src = data_dir / name
        if not src.is_file():
            continue
        target_data.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, target_data / name)
        copied_data += 1

    manifest_path = _write_manifest(
        target_dir=target_dir,
        target_payload=target_payload,
        target_data=target_data,
        version=version,
        changes=changes,
        mandatory=mandatory,
        released=released,
    )
    print(f"Payload: {target_payload} ({target_payload.stat().st_size} bytes)")
    print(f"Data files: {copied_data} in {target_data}")
    return manifest_path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Опубликовать обновление ProtocolOOT на шару")
    parser.add_argument("--exe", type=Path, default=None, help="Путь к собранному ProtocolOOT.exe")
    parser.add_argument("--version", required=True, help="Версия релиза, напр. 1.5.2")
    parser.add_argument(
        "--share-root",
        type=Path,
        required=True,
        help=r"Корень шары, напр. \\SERVER\SOFT\ProtocolOOT",
    )
    parser.add_argument(
        "--data-dir",
        type=Path,
        default=None,
        help="Каталог data/ для публикации (по умолчанию — data/ рядом с exe)",
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
    parser.add_argument(
        "--refresh-manifest",
        action="store_true",
        help="Только пересобрать manifest.json в windows/<версия>/ (exe и data/ уже на шаре)",
    )
    args = parser.parse_args(argv)

    share_root = args.share_root.expanduser().resolve()
    version = args.version.strip()
    changes = [str(c) for c in args.changes if str(c).strip()]

    if args.refresh_manifest:
        manifest_path = refresh_manifest(
            version_dir=share_root / "windows" / version,
            version=version,
            changes=changes or [f"Обновление {version}"],
            mandatory=bool(args.mandatory),
            released=str(args.released).strip(),
        )
        print(f"Manifest: {manifest_path}")
        return 0

    if args.exe is None:
        parser.error("--exe обязателен без --refresh-manifest")

    manifest_path = publish(
        exe_path=args.exe.expanduser().resolve(),
        version=version,
        share_root=share_root,
        changes=changes,
        mandatory=bool(args.mandatory),
        released=str(args.released).strip(),
        data_src_dir=args.data_dir.expanduser().resolve() if args.data_dir else None,
    )
    print(f"Manifest: {manifest_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
