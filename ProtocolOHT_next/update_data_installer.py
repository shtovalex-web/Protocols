# -*- coding: utf-8
"""Установка файлов data/ при автообновлении."""

from __future__ import annotations

import os
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path

from update_bundle_files import DATA_SUBDIR_NAME
from update_installer import UpdateInstallerError, stage_payload_copy
from update_info import write_update_info
from update_manifest import DataFilePayload, UpdateManifest

LEGACY_DATA_FILE_BACKUP_SUFFIX = ".bak"


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


def _cleanup_legacy_bak_files(data_dir: Path) -> None:
    """Удаляет устаревшие *.bak из data/ (от прежней схемы отката)."""
    if not data_dir.is_dir():
        return
    for path in data_dir.glob(f"*{LEGACY_DATA_FILE_BACKUP_SUFFIX}"):
        if not path.is_file():
            continue
        try:
            path.unlink()
        except OSError:
            pass


def _backup_data_file(temp_dir: Path, destination: Path) -> Path | None:
    """Копия файла во временный каталог; None — файла не было."""
    if not destination.is_file():
        return None
    suffix = destination.suffix if destination.suffix else ".bak"
    handle, temp_name = tempfile.mkstemp(
        prefix=f"{destination.stem}_",
        suffix=suffix,
        dir=temp_dir,
    )
    os.close(handle)
    backup = Path(temp_name)
    shutil.copy2(destination, backup)
    return backup


def _discard_data_file_backups(rollbacks: list[_ReplacedFileRollback]) -> None:
    for item in rollbacks:
        if item.backup and item.backup.is_file():
            try:
                item.backup.unlink()
            except OSError:
                pass


def _restore_data_file_backups(rollbacks: list[_ReplacedFileRollback]) -> None:
    for item in rollbacks:
        if item.backup and item.backup.is_file():
            try:
                shutil.copy2(item.backup, item.destination)
            except OSError:
                pass
            try:
                item.backup.unlink()
            except OSError:
                pass
        elif item.destination.is_file():
            try:
                item.destination.unlink()
            except OSError:
                pass


def _permission_error_hint(path: Path) -> UpdateInstallerError:
    name = path.name.lower()
    if name.endswith(".docx"):
        hint = "Закройте документы Word (шаблоны протокола, инструкции) и повторите обновление."
    else:
        hint = "Закройте файлы из папки data/ и повторите обновление."
    return UpdateInstallerError(f"Нет доступа к файлу:\n{path}\n\n{hint}")


def _stage_data_file_copy(
    source: Path,
    destination: Path,
    *,
    expected_sha256: str,
    expected_size: int,
) -> None:
    try:
        stage_payload_copy(
            source,
            destination,
            expected_sha256=expected_sha256,
            expected_size=expected_size,
        )
    except OSError as error:
        if isinstance(error, PermissionError) or getattr(error, "winerror", None) == 5:
            raise _permission_error_hint(destination) from error
        raise


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

    data_dir = data_dir_for_exe(exe_path)
    data_dir.mkdir(parents=True, exist_ok=True)
    _cleanup_legacy_bak_files(data_dir)

    with tempfile.TemporaryDirectory(prefix="ProtocolOOT_data_update_") as temp_raw:
        temp_dir = Path(temp_raw)
        rollbacks: list[_ReplacedFileRollback] = []

        try:
            for entry in entries:
                source = data_file_source(manifest_path, entry)
                destination = data_file_destination(exe_path, entry)
                try:
                    backup = _backup_data_file(temp_dir, destination)
                except OSError as error:
                    if isinstance(error, PermissionError) or getattr(error, "winerror", None) == 5:
                        raise _permission_error_hint(destination) from error
                    raise
                rollbacks.append(_ReplacedFileRollback(destination, backup))
                _stage_data_file_copy(
                    source,
                    destination,
                    expected_sha256=entry.sha256,
                    expected_size=entry.size,
                )
            _discard_data_file_backups(rollbacks)
        except (UpdateInstallerError, OSError):
            _restore_data_file_backups(rollbacks)
            raise
