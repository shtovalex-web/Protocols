# -*- coding: utf-8
"""Обратная совместимость: windows_app_bundle → desktop_updater."""

import protocol_updater_config  # noqa: F401

from desktop_updater.registry import get_config
from desktop_updater.windows_app_bundle import (  # noqa: F401
    APP_UPDATE_STAGING_DIR,
    INTERNAL_DIR_NAME,
    app_update_staging_dir,
    create_windows_app_zip,
    is_windows_app_bundle,
    resolve_windows_payload,
    staged_app_zip_path,
)

_cfg = get_config()
WINDOWS_APP_ZIP_NAME = _cfg.resolved_app_bundle_zip_name
DEFAULT_EXE_NAME = _cfg.exe_name
