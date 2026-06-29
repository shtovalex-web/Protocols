# -*- coding: utf-8
"""Обратная совместимость: update_info → desktop_updater."""

import protocol_updater_config  # noqa: F401

from desktop_updater.info import (  # noqa: F401
    UpdateInfo,
    data_dir_for_exe,
    installed_version_from_data,
    load_update_info,
    update_info_path,
    write_update_info,
)
from protocol_updater_config import DATA_SUBDIR_NAME  # noqa: F401
from desktop_updater.info import UPDATE_INFO_FILENAME  # noqa: F401
