# -*- coding: utf-8
"""Конфигурация пакета desktop_updater для конкретного приложения."""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class UpdaterConfig:
    """Параметры автообновления, задаваемые один раз при старте приложения."""

    app_name: str
    exe_name: str
    app_version: str
    default_share_root: Path
    env_prefix: str = "DESKTOP_APP"
    linux_exe_name: str | None = None
    update_config_filename: str = "update_config.json"
    data_subdir: str = "data"
    data_replace_filenames: tuple[str, ...] = field(default_factory=tuple)
    user_dir_resolver: Callable[[], Path] | None = None
    restart_cmd_name: str | None = None
    app_bundle_zip_name: str | None = None
    data_update_temp_prefix: str = "desktop_updater_data_"
    publish_script_hint: str = "tools/publish_update_manifest.py"
    platform_subdirs: tuple[str, ...] = ("windows", "linux")

    def __post_init__(self) -> None:
        if self.linux_exe_name is None:
            self.linux_exe_name = Path(self.exe_name).stem

    @property
    def env_manifest(self) -> str:
        return f"{self.env_prefix}_UPDATE_MANIFEST"

    @property
    def env_force_check(self) -> str:
        return f"{self.env_prefix}_UPDATE_CHECK"

    @property
    def default_manifest_path(self) -> Path:
        return self.default_share_root / "manifest.json"

    @property
    def resolved_restart_cmd_name(self) -> str:
        if self.restart_cmd_name:
            return self.restart_cmd_name
        stem = Path(self.exe_name).stem
        return f"{stem}_restart.cmd"

    @property
    def resolved_app_bundle_zip_name(self) -> str:
        if self.app_bundle_zip_name:
            return self.app_bundle_zip_name
        stem = Path(self.exe_name).stem
        return f"{stem}_app.zip"

    def user_dir(self) -> Path:
        if self.user_dir_resolver is not None:
            return self.user_dir_resolver()
        return Path.cwd()

    def platform_exe_names(self) -> dict[str, str]:
        return {
            "windows": self.exe_name,
            "linux": self.linux_exe_name or Path(self.exe_name).stem,
        }
