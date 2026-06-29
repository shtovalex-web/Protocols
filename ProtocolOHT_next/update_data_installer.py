# -*- coding: utf-8
"""Установка файлов data/ при автообновлении."""

from __future__ import annotations

import shutil
from dataclasses import dataclass
from pathlib import Path

from update_bundle_files import DATA_SUBDIR_NAME
from update_installer import UpdateInstallerError, stage_payload_copy
from update_info import write_update_info
from update_manifest import DataFilePayload, UpdateManifest

DATA_FILE_BACKUP_SUFFIX = ".bak"


@dataclass
class _ReplacedFileRollback:
    destination: Path
    backup: Path | None


def data_dir_for_exe(exe_path: Path) -> Path:
    return exe_path.resolve().parent / DATA_SUBDIR_NAME


def data_file_source(manifest_path: Path, entry: DataFilePayload) -> Path:
    return manifest_path.parent / entry.relative_path


def data_file_destination(exe_path: Path, entry: DataFilePayload) -> Path:
    name = Path(entry.relative_path.replace("\\", "/")).name
    return data_dir_for_exe(exe_path) / name


def _backup_data_file(destination: Path) -> Path | None:
    """Копия одного файла как <имя>.bak; None — файла не было."""
    if not destination.is_file():
        return None
    backup = destination.with_name(destination.name + DATA_FILE_BACKUP_SUFFIX)
    if backup.is_file():
        backup.unlink()
    shutil.copy2(destination, backup)
    return backup


def _discard_data_file_backups(rollbacks: list[_ReplacedFileRollback]) -> None:
    for item in rollbacks:
        if item.backup and item.backup.is_file():
            item.backup.unlink()


def _restore_data_file_backups(rollbacks: list[_ReplacedFileRollback]) -> None:
    for item in rollbacks:
        if item.backup and item.backup.is_file():
            shutil.copy2(item.backup, item.destination)
            item.backup.unlink()
        elif item.destination.is_file():
            item.destination.unlink()


def mark_data_version_installed(exe_path: Path, manifest: UpdateManifest) -> None:
    """Обновляет data/update_info.json после успешной установки релиза."""
    write_update_info(
        data_dir_for_exe(exe_path),
        version=manifest.latest_version,
        released=manifest.released or None,
    )


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
    rollbacks: list[_ReplacedFileRollback] = []

    try:
        for entry in entries:
            source = data_file_source(manifest_path, entry)
            destination = data_file_destination(exe_path, entry)
            backup = _backup_data_file(destination)
            rollbacks.append(_ReplacedFileRollback(destination, backup))
            stage_payload_copy(
                source,
                destination,
                expected_sha256=entry.sha256,
                expected_size=entry.size,
            )
        _discard_data_file_backups(rollbacks)
    except (UpdateInstallerError, OSError):
        _restore_data_file_backups(rollbacks)
        raise
