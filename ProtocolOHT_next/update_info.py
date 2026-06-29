# -*- coding: utf-8
"""Маркер установленной версии комплекта data/ (update_info.json)."""

from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from update_bundle_files import DATA_SUBDIR_NAME

UPDATE_INFO_FILENAME = "update_info.json"


@dataclass
class UpdateInfo:
    version: str
    released: str = ""
    platform: str = "windows"

    @classmethod
    def from_dict(cls, data: object) -> UpdateInfo | None:
        if not isinstance(data, dict):
            return None
        version = str(data.get("version", "")).strip()
        if not version:
            return None
        released = str(data.get("released", "") or "").strip()
        platform = str(data.get("platform", "windows") or "windows").strip() or "windows"
        return cls(version=version, released=released, platform=platform)

    def to_dict(self) -> dict[str, str]:
        payload: dict[str, str] = {"version": self.version, "platform": self.platform}
        if self.released:
            payload["released"] = self.released
        return payload


def update_info_path(data_dir: Path) -> Path:
    return data_dir / UPDATE_INFO_FILENAME


def data_dir_for_exe(exe_path: Path) -> Path:
    return exe_path.resolve().parent / DATA_SUBDIR_NAME


def write_update_info(
    data_dir: Path,
    *,
    version: str,
    released: str | None = None,
    platform: str = "windows",
) -> Path:
    """Записывает data/update_info.json; возвращает путь к файлу."""
    info = UpdateInfo(
        version=version.strip(),
        released=(released or date.today().isoformat()).strip(),
        platform=platform.strip() or "windows",
    )
    path = update_info_path(data_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(info.to_dict(), ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return path


def load_update_info(data_dir: Path) -> UpdateInfo | None:
    path = update_info_path(data_dir)
    if not path.is_file():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return UpdateInfo.from_dict(data)
    except (OSError, json.JSONDecodeError):
        return None


def installed_version_from_data(exe_path: Path) -> str | None:
    info = load_update_info(data_dir_for_exe(exe_path))
    if info is None:
        return None
    return info.version
