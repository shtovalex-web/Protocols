# -*- coding: utf-8
"""Обратная совместимость: update_installer → desktop_updater."""

import subprocess
import sys

import protocol_updater_config  # noqa: F401

from desktop_updater.installer import (  # noqa: F401
    UpdateInstallerError,
    apply_pending_app_staging,
    backup_exe_path,
    cleanup_app_update_staging,
    cleanup_backup_exe,
    cleanup_restart_cmd,
    exit_after_update,
    exit_for_update_restart,
    launch_updated_exe,
    restart_cmd_path,
    stage_payload_copy,
    stage_windows_app_bundle,
    staged_new_exe_path,
    swap_exe_via_rename,
    _launch_updated_exe_cmd_helper,
    _launch_updated_exe_subprocess,
    _launch_updated_exe_windows,
    _write_restart_cmd,
)
