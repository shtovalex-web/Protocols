# -*- coding: utf-8 -*-
"""Windows onedir: exe + _internal, доставка через ProtocolOOT_app.zip."""

from __future__ import annotations

import zipfile
from pathlib import Path

WINDOWS_APP_ZIP_NAME = "ProtocolOOT_app.zip"
APP_UPDATE_STAGING_DIR = ".app_update_staging"
INTERNAL_DIR_NAME = "_internal"
DEFAULT_EXE_NAME = "ProtocolOOT.exe"


def is_windows_app_bundle(path: Path) -> bool:
    name = path.name.lower()
    return name == WINDOWS_APP_ZIP_NAME.lower() or name.endswith(".zip")


def staged_app_zip_path(install_dir: Path) -> Path:
    return install_dir.resolve() / f"{WINDOWS_APP_ZIP_NAME}.new"


def app_update_staging_dir(install_dir: Path) -> Path:
    return install_dir.resolve() / APP_UPDATE_STAGING_DIR


def create_windows_app_zip(install_dir: Path, *, exe_name: str = DEFAULT_EXE_NAME) -> Path:
    root = install_dir.resolve()
    exe = root / exe_name
    internal = root / INTERNAL_DIR_NAME
    if not exe.is_file():
        msg = f"EXE not found: {exe}"
        raise FileNotFoundError(msg)
    if not internal.is_dir():
        msg = f"Missing runtime folder: {internal}"
        raise FileNotFoundError(msg)
    zip_path = root / WINDOWS_APP_ZIP_NAME
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.write(exe, exe.name)
        for item in internal.rglob("*"):
            if item.is_file():
                arcname = Path(INTERNAL_DIR_NAME) / item.relative_to(internal)
                archive.write(item, arcname.as_posix())
    return zip_path


def resolve_windows_payload(version_dir: Path) -> Path:
    zip_path = version_dir / WINDOWS_APP_ZIP_NAME
    if zip_path.is_file():
        return zip_path
    return version_dir / DEFAULT_EXE_NAME
