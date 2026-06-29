# -*- coding: utf-8
"""Переиспользуемый модуль автообновления desktop .exe по manifest на сетевой шаре."""

from desktop_updater.config import UpdaterConfig
from desktop_updater.registry import configure, get_config, try_get_config

__all__ = [
    "UpdaterConfig",
    "configure",
    "get_config",
    "try_get_config",
]
