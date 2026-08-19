# -*- coding: utf-8
"""Обратная совместимость: startup_update → desktop_updater."""

import protocol_updater_config  # noqa: F401

from desktop_updater.registry import get_config
from desktop_updater.startup import (  # noqa: F401
    app_version,
    check_updates_interactive,
    current_exe_path,
    is_frozen,
    parse_changelog_version,
    prepare_startup_updates,
    should_check_for_updates,
)

ENV_FORCE_CHECK = get_config().env_force_check
