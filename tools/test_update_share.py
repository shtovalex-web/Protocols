# -*- coding: utf-8
"""Проверка тестовой шары обновлений (без GUI)."""

from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
_NEXT = ROOT / "ProtocolOHT_next"
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(_NEXT))
sys.path.insert(0, str(ROOT / "tools"))

from protocol_app_info import APP_VERSION  # noqa: E402
from startup_update import app_version  # noqa: E402
from update_config import resolve_update_share_root  # noqa: E402
from update_installer import stage_payload_copy  # noqa: E402
from update_scan import resolve_latest_update  # noqa: E402
from publish_update_manifest import verify_manifest_payload  # noqa: E402

SHARE_ROOT = Path(r"D:\Обновление")


def main() -> int:
    if not SHARE_ROOT.is_dir():
        print(f"FAIL: нет каталога шары: {SHARE_ROOT}")
        return 1

    current = app_version() or (APP_VERSION or "").strip()
    resolved = resolve_latest_update(SHARE_ROOT, current_version=current, platform="windows")
    if resolved is None:
        print(f"Текущая версия: {current}")
        print(f"Каталог шары: {SHARE_ROOT}")
        print("OK: актуальная версия или нет релизов в windows/<версия>/")
        return 0

    manifest_path = resolved.anchor_manifest_path
    manifest = resolved.manifest
    payload = manifest.windows_payload_path(manifest_path)

    print(f"Текущая версия: {current}")
    print(f"Каталог шары: {resolve_update_share_root(SHARE_ROOT)}")
    print(f"Релиз на шаре: {resolved.version}")
    print(f"Манифест: {manifest_path}")
    print(f"Файл обновления: {payload}")

    errors = verify_manifest_payload(manifest_path)
    if errors:
        print("FAIL: manifest.json не совпадает с файлами на шаре:")
        for item in errors:
            print(f"  - {item}")
        print(
            "Исправление: py -3 tools/publish_update_manifest.py "
            '--exe "…\\ProtocolOOT.exe" --version … --share-root "D:\\Обновление"'
        )
        return 1

    if not payload.is_file():
        print(f"FAIL: нет файла обновления: {payload}")
        return 1

    staged_parent = manifest_path.parent / "_verify_staging"
    staged_parent.mkdir(exist_ok=True)
    staged = staged_parent / "ProtocolOOT.exe.new"
    try:
        stage_payload_copy(
            payload,
            staged,
            expected_sha256=manifest.windows.sha256,
            expected_size=manifest.windows.size,
        )
        print(f"OK: exe — копия и sha256 ({staged.stat().st_size} байт)")
    finally:
        staged.unlink(missing_ok=True)

    config_path = ROOT / "ProtocolOHT_onefile" / "update_config.json"
    if config_path.is_file():
        cfg = json.loads(config_path.read_text(encoding="utf-8"))
        print(f"OK: update_config.json -> {cfg.get('manifest_path')}")
    else:
        print(f"WARN: нет {config_path}")

    print("Готово: запустите ProtocolOOT.exe — должно предложить обновление.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
