# -*- coding: utf-8
"""Windows onedir: exe + _internal, доставка через zip (legacy manifest)."""

from __future__ import annotations

import zipfile
from pathlib import Path

from desktop_updater.registry import get_config

APP_UPDATE_STAGING_DIR = ".app_update_staging"
INTERNAL_DIR_NAME = "_internal"


def windows_app_zip_name() -> str:
    return get_config().resolved_app_bundle_zip_name


def default_exe_name() -> str:
    return get_config().exe_name


def is_windows_app_bundle(path: Path) -> bool:
    zip_name = windows_app_zip_name().lower()
    name = path.name.lower()
    return name == zip_name or name.endswith(".zip")


def staged_app_zip_path(install_dir: Path) -> Path:
    return install_dir.resolve() / f"{windows_app_zip_name()}.new"


def app_update_staging_dir(install_dir: Path) -> Path:
    return install_dir.resolve() / APP_UPDATE_STAGING_DIR


def create_windows_app_zip(install_dir: Path, *, exe_name: str | None = None) -> Path:
    cfg = get_config()
    root = install_dir.resolve()
    exe_file = exe_name or cfg.exe_name
    exe = root / exe_file
    internal = root / INTERNAL_DIR_NAME
    if not exe.is_file():
        msg = f"EXE not found: {exe}"
        raise FileNotFoundError(msg)
    if not internal.is_dir():
        msg = f"Missing runtime folder: {internal}"
        raise FileNotFoundError(msg)
    zip_path = root / cfg.resolved_app_bundle_zip_name
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.write(exe, exe.name)
        for item in internal.rglob("*"):
            if item.is_file():
                arcname = Path(INTERNAL_DIR_NAME) / item.relative_to(internal)
                archive.write(item, arcname.as_posix())
    return zip_path


def resolve_windows_payload(version_dir: Path) -> Path:
    cfg = get_config()
    zip_path = version_dir / cfg.resolved_app_bundle_zip_name
    if zip_path.is_file():
        return zip_path
    return version_dir / cfg.exe_name
