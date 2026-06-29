# -*- coding: utf-8
"""Обратная совместимость: pending_changelog → desktop_updater."""

import protocol_updater_config  # noqa: F401

from desktop_updater.pending_changelog import (  # noqa: F401
    PENDING_FILENAME,
    pending_changelog_path,
    pop_pending_changelog,
    write_pending_changelog,
)
