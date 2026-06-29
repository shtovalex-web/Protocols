# -*- coding: utf-8
"""Обратная совместимость: update_bundle_files → desktop_updater."""

import protocol_updater_config  # noqa: F401

from desktop_updater.bundle_files import (  # noqa: F401
    DATA_POLICY_REPLACE,
    build_data_manifest_entries,
)
from protocol_updater_config import DATA_REPLACE_FILENAMES, DATA_SUBDIR_NAME  # noqa: F401
