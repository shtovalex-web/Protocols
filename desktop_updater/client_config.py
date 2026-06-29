# -*- coding: utf-8
"""Локальная конфигурация проверки обновлений (update_config.json)."""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
from pathlib import Path

from desktop_updater.registry import get_config
from desktop_updater.version_compare import parse_version

_MANIFEST_PATH_RE = re.compile(r'"manifest_path"\s*:\s*"([^"]*)"')
_ENABLED_RE = re.compile(r'"enabled"\s*:\s*(true|false)', re.IGNORECASE)


@dataclass
class UpdateConfig:
    manifest_path: Path | None = None
    enabled: bool = True

    def __post_init__(self) -> None:
        if self.manifest_path is None:
            self.manifest_path = get_config().default_share_root


def resolve_update_share_root(config_path: Path) -> Path:
    """Каталог шары: из пути к manifest.json, к папке версии или напрямую из каталога."""
    resolved = config_path.expanduser().resolve()
    if resolved.is_dir():
        return resolved
    if resolved.name.lower() != "manifest.json":
        return resolved.parent
    parent = resolved.parent
    try:
        parse_version(parent.name)
    except ValueError:
        return parent
    if parent.parent.name.lower() == "windows":
        return parent.parent.parent
    return parent


def update_config_path() -> Path:
    cfg = get_config()
    return cfg.user_dir() / cfg.update_config_filename


def format_manifest_path_for_json(path: Path | str) -> str:
    """Путь для JSON: прямые слэши (безопасно при ручном редактировании на Windows)."""
    return str(path).replace("\\", "/")


def _config_from_dict(data: dict) -> UpdateConfig:
    cfg = get_config()
    manifest_raw = data.get("manifest_path", cfg.default_manifest_path)
    enabled = bool(data.get("enabled", True))
    return UpdateConfig(manifest_path=Path(str(manifest_raw)), enabled=enabled)


def parse_update_config_text(raw: str) -> UpdateConfig | None:
    """Разбор текста конфигурации; None — не удалось прочитать."""
    text = (raw or "").strip()
    if not text:
        return None

    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        manifest_match = _MANIFEST_PATH_RE.search(text)
        if manifest_match is None:
            return None
        enabled = True
        enabled_match = _ENABLED_RE.search(text)
        if enabled_match is not None:
            enabled = enabled_match.group(1).lower() == "true"
        return UpdateConfig(
            manifest_path=Path(manifest_match.group(1)),
            enabled=enabled,
        )

    if not isinstance(data, dict):
        return None
    return _config_from_dict(data)


def load_update_config() -> UpdateConfig:
    cfg = get_config()
    env_path = os.environ.get(cfg.env_manifest, "").strip()
    if env_path:
        return UpdateConfig(manifest_path=Path(env_path))

    path = update_config_path()
    if not path.is_file():
        return UpdateConfig()

    try:
        raw = path.read_text(encoding="utf-8")
    except OSError:
        return UpdateConfig()

    parsed = parse_update_config_text(raw)
    if parsed is None:
        return UpdateConfig()
    return parsed
