# -*- coding: utf-8
"""Просмотр доступных обновлений на шаре (сканирование вложенных каталогов)."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
_NEXT = ROOT / "ProtocolOHT_next"
sys.path.insert(0, str(_NEXT))
sys.path.insert(1, str(ROOT))

from protocol_app_info import APP_VERSION  # noqa: E402
from update_config import load_update_config, resolve_update_share_root  # noqa: E402
from update_scan import resolve_latest_update, scan_update_candidates  # noqa: E402
from version_compare import is_newer_version  # noqa: E402


def main() -> int:
    config = load_update_config()
    share_root = resolve_update_share_root(config.manifest_path)
    current = (APP_VERSION or "").strip()

    print(f"Текущая версия: {current}")
    print(f"Каталог шары: {share_root}")
    print()

    if not share_root.is_dir():
        print(f"FAIL: каталог недоступен: {share_root}")
        return 1

    candidates = scan_update_candidates(share_root)
    if not candidates:
        print("На шаре не найдено ни одного обновления (manifest.json / ProtocolOOT.exe).")
        return 1

    print("Найденные версии:")
    for item in candidates:
        flag = "NEW" if is_newer_version(item.version, current) else "   "
        payload = item.manifest.windows_payload_path(item.anchor_manifest_path)
        print(f"  [{flag}] {item.version} — {payload} ({item.source})")

    resolved = resolve_latest_update(share_root, current_version=current)
    print()
    if resolved is None:
        print("Установлена актуальная версия — обновление не требуется.")
        return 0

    print(f"Рекомендуемое обновление: {resolved.version}")
    print(f"Файл: {resolved.manifest.windows_payload_path(resolved.anchor_manifest_path)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
