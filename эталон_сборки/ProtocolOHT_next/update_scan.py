# -*- coding: utf-8
"""Обратная совместимость: update_scan → desktop_updater."""

import protocol_updater_config  # noqa: F401

from desktop_updater.registry import get_config
from desktop_updater.scan import (  # noqa: F401
    UpdateCandidate,
    default_update_platform,
    manifest_from_exe,
    pick_newest_update,
    resolve_latest_update,
    scan_update_candidates,
    share_root_from_manifest,
    version_from_dir_name,
)

_cfg = get_config()
DEFAULT_WINDOWS_EXE_NAME = _cfg.exe_name
DEFAULT_LINUX_EXE_NAME = _cfg.linux_exe_name or _cfg.exe_name
UPDATE_PLATFORM_SUBDIRS = _cfg.platform_subdirs
PLATFORM_EXE_NAMES = _cfg.platform_exe_names()
