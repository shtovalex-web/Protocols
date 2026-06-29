# -*- coding: utf-8
"""Обратная совместимость: update_config → desktop_updater."""

import protocol_updater_config  # noqa: F401

from desktop_updater.client_config import (  # noqa: F401
    UpdateConfig,
    format_manifest_path_for_json,
    load_update_config,
    parse_update_config_text,
    resolve_update_share_root,
    update_config_path,
)
from desktop_updater.registry import get_config

_cfg = get_config()
DEFAULT_UPDATE_SHARE_ROOT = _cfg.default_share_root
DEFAULT_MANIFEST_PATH = _cfg.default_manifest_path
UPDATE_CONFIG_FILENAME = _cfg.update_config_filename
ENV_MANIFEST = _cfg.env_manifest
ENV_FORCE_CHECK = _cfg.env_force_check
