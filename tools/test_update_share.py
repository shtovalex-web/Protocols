# -*- coding: utf-8 -*-
"""Проверка тестовой шары обновлений (без GUI)."""

from __future__ import annotations

import json
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
_NEXT = ROOT / "ProtocolOHT_next"
sys.path.insert(0, str(_NEXT))

from protocol_app_info import APP_VERSION  # noqa: E402
from update_installer import stage_payload_copy  # noqa: E402
from update_manifest import load_update_manifest  # noqa: E402
from version_compare import is_newer_version  # noqa: E402

SHARE_ROOT = Path(r"D:\Обновление")
MANIFEST_PATH = SHARE_ROOT / "manifest.json"


def main() -> int:
    if not MANIFEST_PATH.is_file():
        print(f"FAIL: нет манифеста: {MANIFEST_PATH}")
        return 1

    manifest = load_update_manifest(MANIFEST_PATH)
    payload = manifest.windows_payload_path(MANIFEST_PATH)
    current = (APP_VERSION or "").strip()

    print(f"Текущая версия в исходниках: {current}")
    print(f"Версия в манифесте: {manifest.latest_version}")
    print(f"Файл обновления: {payload}")

    if not payload.is_file():
        print(f"FAIL: нет файла обновления: {payload}")
        return 1

    payload_size = payload.stat().st_size
    if payload_size != manifest.windows.size:
        print(
            f"FAIL: размер {payload} ({payload_size}) != manifest.json ({manifest.windows.size})"
        )
        print("Исправление: py -3 tools/publish_update_manifest.py --exe ... --share-root D:\\Обновление")
        return 1

    if not is_newer_version(manifest.latest_version, current):
        print("FAIL: манифест не новее текущей версии — обновление не предложится")
        return 1

    with tempfile.TemporaryDirectory() as tmp:
        staged = Path(tmp) / "ProtocolOOT.exe.new"
        stage_payload_copy(
            payload,
            staged,
            expected_sha256=manifest.windows.sha256,
            expected_size=manifest.windows.size,
        )
        print(f"OK: копия и sha256 проверены ({staged.stat().st_size} байт)")

    config_path = ROOT / "ProtocolOHT_onefile" / "update_config.json"
    if config_path.is_file():
        cfg = json.loads(config_path.read_text(encoding="utf-8"))
        print(f"OK: update_config.json -> {cfg.get('manifest_path')}")
    else:
        print(f"WARN: нет {config_path}")

    print("Готово: запустите ProtocolOHT_onefile\\ProtocolOOT.exe — должно предложить обновление.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
