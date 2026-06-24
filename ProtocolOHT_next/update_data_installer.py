# -*- coding: utf-8
"""Установка файлов data/ при автообновлении."""

from __future__ import annotations

import shutil
from pathlib import Path

from update_bundle_files import DATA_POLICY_REPLACE, DATA_SUBDIR_NAME
from update_installer import UpdateInstallerError, stage_payload_copy
from update_manifest import DataFilePayload, UpdateManifest


def data_dir_for_exe(exe_path: Path) -> Path:
    return exe_path.resolve().parent / DATA_SUBDIR_NAME


def data_backup_dir_for_exe(exe_path: Path) -> Path:
    return exe_path.resolve().parent / f"{DATA_SUBDIR_NAME}.backup"


def data_file_source(manifest_path: Path, entry: DataFilePayload) -> Path:
    return manifest_path.parent / entry.relative_path


def data_file_destination(exe_path: Path, entry: DataFilePayload) -> Path:
    name = Path(entry.relative_path.replace("\\", "/")).name
    return data_dir_for_exe(exe_path) / name


def _backup_data_dir(exe_path: Path) -> None:
    src = data_dir_for_exe(exe_path)
    if not src.is_dir():
        return
    backup = data_backup_dir_for_exe(exe_path)
    if backup.exists():
        shutil.rmtree(backup)
    shutil.copytree(src, backup)


def _restore_data_backup(exe_path: Path) -> None:
    backup = data_backup_dir_for_exe(exe_path)
    if not backup.is_dir():
        return
    dest = data_dir_for_exe(exe_path)
    if dest.exists():
        shutil.rmtree(dest)
    shutil.copytree(backup, dest)


def apply_data_updates(
    manifest_path: Path,
    manifest: UpdateManifest,
    exe_path: Path,
) -> None:
    """Копирует файлы data/ с шары (policy=replace). Корень exe не затрагивается."""
    entries = manifest.replace_data_files()
    if not entries:
        return

    data_dir_for_exe(exe_path).mkdir(parents=True, exist_ok=True)
    _backup_data_dir(exe_path)

    try:
        for entry in entries:
            source = data_file_source(manifest_path, entry)
            destination = data_file_destination(exe_path, entry)
            stage_payload_copy(
                source,
                destination,
                expected_sha256=entry.sha256,
                expected_size=entry.size,
            )
    except (UpdateInstallerError, OSError):
        _restore_data_backup(exe_path)
        raise
