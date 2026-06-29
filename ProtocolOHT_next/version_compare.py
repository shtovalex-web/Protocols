# -*- coding: utf-8
"""Обратная совместимость: version_compare → desktop_updater."""

import protocol_updater_config  # noqa: F401

from desktop_updater.version_compare import is_newer_version, parse_version  # noqa: F401
