# -*- coding: utf-8
"""Глобальная регистрация UpdaterConfig для текущего процесса."""

from __future__ import annotations

from desktop_updater.config import UpdaterConfig

_CONFIG: UpdaterConfig | None = None


def configure(config: UpdaterConfig) -> UpdaterConfig:
    global _CONFIG
    _CONFIG = config
    return config


def get_config() -> UpdaterConfig:
    if _CONFIG is None:
        msg = "desktop_updater is not configured; call configure(UpdaterConfig(...)) first"
        raise RuntimeError(msg)
    return _CONFIG


def try_get_config() -> UpdaterConfig | None:
    return _CONFIG
