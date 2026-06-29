# -*- coding: utf-8 -*-
"""Отложенный показ «Что нового» после перезапуска (без второго Tk до главного окна)."""

from __future__ import annotations

import json
from pathlib import Path

PENDING_FILENAME = ".pending_changelog.json"


def pending_changelog_path(data_dir: Path) -> Path:
    return data_dir / PENDING_FILENAME


def write_pending_changelog(
    data_dir: Path,
    *,
    version: str,
    changes: list[str],
) -> Path:
    path = pending_changelog_path(data_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {"version": version.strip(), "changes": list(changes)}
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return path


def pop_pending_changelog(data_dir: Path) -> tuple[str, list[str]] | None:
    path = pending_changelog_path(data_dir)
    if not path.is_file():
        return None
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
        path.unlink(missing_ok=True)
    except (OSError, json.JSONDecodeError):
        path.unlink(missing_ok=True)
        return None
    if not isinstance(raw, dict):
        return None
    version = str(raw.get("version", "")).strip()
    if not version:
        return None
    changes_raw = raw.get("changes", [])
    if not isinstance(changes_raw, list):
        changes_raw = []
    changes = [str(item).strip() for item in changes_raw if str(item).strip()]
    return version, changes
