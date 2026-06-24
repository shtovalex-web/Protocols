# -*- coding: utf-8 -*-
"""Локальная конфигурация проверки обновлений."""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path

from app_paths import application_user_dir

DEFAULT_MANIFEST_PATH = Path(r"\\SERVER\SOFT\ProtocolOOT\manifest.json")
UPDATE_CONFIG_FILENAME = "update_config.json"
ENV_MANIFEST = "PROTOCOLOOT_UPDATE_MANIFEST"
ENV_FORCE_CHECK = "PROTOCOLOOT_UPDATE_CHECK"


@dataclass
class UpdateConfig:
    manifest_path: Path = DEFAULT_MANIFEST_PATH
    enabled: bool = True


def update_config_path() -> Path:
    return application_user_dir() / UPDATE_CONFIG_FILENAME


def load_update_config() -> UpdateConfig:
    env_path = os.environ.get(ENV_MANIFEST, "").strip()
    if env_path:
        return UpdateConfig(manifest_path=Path(env_path))

    path = update_config_path()
    if not path.is_file():
        return UpdateConfig()

    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        return UpdateConfig()

    manifest_raw = data.get("manifest_path", DEFAULT_MANIFEST_PATH)
    enabled = bool(data.get("enabled", True))
    return UpdateConfig(manifest_path=Path(str(manifest_raw)), enabled=enabled)
