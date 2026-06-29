# -*- coding: utf-8
"""Обратная совместимость: update_data_installer → desktop_updater."""

import protocol_updater_config  # noqa: F401

from desktop_updater.data_installer import (  # noqa: F401
    LEGACY_DATA_FILE_BACKUP_SUFFIX,
    apply_data_updates,
    data_file_destination,
    data_file_source,
    mark_data_version_installed,
)
from desktop_updater.info import data_dir_for_exe  # noqa: F401
