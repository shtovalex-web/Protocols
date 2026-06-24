# -*- coding: utf-8 -*-
"""Установка обновления: копия с шары и замена .exe через rename."""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

from update_manifest import sha256_file


class UpdateInstallerError(Exception):
    """Ошибка установки обновления."""


def staged_new_exe_path(exe_path: Path) -> Path:
    return exe_path.with_name(f"{exe_path.name}.new")


def backup_exe_path(exe_path: Path) -> Path:
    return exe_path.with_name(f"{exe_path.name}.old")


def stage_payload_copy(
    source: Path,
    destination: Path,
    *,
    expected_sha256: str,
    expected_size: int,
) -> Path:
    if not source.is_file():
        msg = f"Update file not found: {source}"
        raise UpdateInstallerError(msg)
    actual_size = source.stat().st_size
    if actual_size != expected_size:
        msg = f"Update file size mismatch: {actual_size} != {expected_size}"
        raise UpdateInstallerError(msg)

    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)
    digest = sha256_file(destination)
    if digest.lower() != expected_sha256.lower():
        destination.unlink(missing_ok=True)
        msg = "Update file checksum mismatch."
        raise UpdateInstallerError(msg)
    return destination


def swap_exe_via_rename(exe_path: Path) -> None:
    new_path = staged_new_exe_path(exe_path)
    old_path = backup_exe_path(exe_path)
    if not new_path.is_file():
        msg = f"Staged update not found: {new_path}"
        raise UpdateInstallerError(msg)

    if old_path.exists():
        old_path.unlink()

    os.replace(exe_path, old_path)
    try:
        os.replace(new_path, exe_path)
    except OSError as error:
        if not exe_path.exists():
            try:
                os.replace(old_path, exe_path)
            except OSError:
                pass
        else:
            exe_path.unlink(missing_ok=True)
            os.replace(old_path, exe_path)
        msg = "Failed to install update."
        raise UpdateInstallerError(msg) from error


def cleanup_backup_exe(exe_path: Path) -> None:
    backup = backup_exe_path(exe_path)
    if backup.is_file():
        backup.unlink()


def launch_updated_exe(exe_path: Path, *, show_changelog: bool, version: str) -> None:
    args = [str(exe_path)]
    if show_changelog:
        args.append(f"--show-changelog={version}")
    subprocess.Popen(args, close_fds=True, cwd=str(exe_path.parent))


def exit_for_update_restart() -> None:
    sys.exit(0)
