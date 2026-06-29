# -*- coding: utf-8
"""Конфигурация desktop_updater для ProtocolOOT."""

from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[1]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from app_paths import application_user_dir
from desktop_updater import UpdaterConfig, configure
from protocol_app_info import APP_VERSION

DATA_SUBDIR_NAME = "data"

DEFAULT_UPDATE_SHARE_UNC = Path(
    r"\\tn.tngrp.ru\df\AK\SHR\Distr_О_О\Шитов Алексей Александрович"
)

DATA_REPLACE_FILENAMES: tuple[str, ...] = (
    "default_protocol.docx",
    "default_protocol_tehnicheskiy.docx",
    "ПОДРОБНАЯ_ИНСТРУКЦИЯ_для_пользователя.docx",
    "ИНСТРУКЦИЯ_оформление_протоколов_Минтруд.docx",
    "ЖУРНАЛ_ДОРАБОТОК.md",
    "FAQ.txt",
    "Шаблон_Минтруд_XSD_УМН.xlsx",
    "!! Шаблон_Минтруд_XSD_УМН _ общ+.xlsx",
    "Шаблон_Минтруд_XSD_УМН _ общ+.xlsx",
    "icon.ico",
)

configure(
    UpdaterConfig(
        app_name="ProtocolOOT",
        exe_name="ProtocolOOT.exe",
        linux_exe_name="ProtocolOOT",
        app_version=APP_VERSION,
        default_share_root=DEFAULT_UPDATE_SHARE_UNC,
        env_prefix="PROTOCOLOOT",
        data_subdir=DATA_SUBDIR_NAME,
        data_replace_filenames=DATA_REPLACE_FILENAMES,
        user_dir_resolver=application_user_dir,
        restart_cmd_name="ProtocolOOT_restart.cmd",
        app_bundle_zip_name="ProtocolOOT_app.zip",
        data_update_temp_prefix="ProtocolOOT_data_update_",
        publish_script_hint="tools/publish_update_manifest.py",
    )
)
