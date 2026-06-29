# -*- coding: utf-8
"""Установка обновления: копия с шары и замена .exe через rename."""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

from update_manifest import sha256_file

# ShellExecuteW: > 32 — успех, иначе код ошибки Windows.
_SHELL_EXECUTE_SUCCESS_MIN = 32
_SW_SHOWNORMAL = 1


class UpdateInstallerError(Exception):
    """Ошибка установки обновления."""


def _safe_unlink(path: Path) -> bool:
    try:
        path.unlink()
        return True
    except OSError:
        return False


def _unlink_with_retries(path: Path, *, attempts: int = 5, delay_sec: float = 0.25) -> bool:
    for attempt in range(attempts):
        if _safe_unlink(path):
            return True
        if attempt + 1 < attempts:
            time.sleep(delay_sec)
    return False


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
        msg = (
            f"Размер файла обновления не совпадает с manifest.json: "
            f"{actual_size} != {expected_size}. "
            f"Пересоберите и опубликуйте обновление: tools/publish_update_manifest.py"
        )
        raise UpdateInstallerError(msg)

    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)
    digest = sha256_file(destination)
    if digest.lower() != expected_sha256.lower():
        destination.unlink(missing_ok=True)
        msg = f"Update file checksum mismatch: {source.name}"
        raise UpdateInstallerError(msg)
    return destination


def swap_exe_via_rename(exe_path: Path) -> None:
    new_path = staged_new_exe_path(exe_path)
    old_path = backup_exe_path(exe_path)
    if not new_path.is_file():
        msg = f"Staged update not found: {new_path}"
        raise UpdateInstallerError(msg)

    if old_path.exists() and not _safe_unlink(old_path):
        stale = exe_path.with_name(f"{exe_path.stem}.old.{os.getpid()}{exe_path.suffix}")
        try:
            os.replace(old_path, stale)
        except OSError:
            pass

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


def cleanup_backup_exe(exe_path: Path) -> bool:
    """Удалить .exe.old после обновления; False — файл ещё занят (не критично)."""
    backup = backup_exe_path(exe_path)
    if not backup.is_file():
        return True
    return _unlink_with_retries(backup)


def _launch_updated_exe_windows(exe_path: Path, *, params: str, cwd: str) -> bool:
    """ShellExecuteW — корректные пути с пробелами и «!»; True при успехе."""
    import ctypes

    result = int(
        ctypes.windll.shell32.ShellExecuteW(
            None,
            "open",
            str(exe_path),
            params or None,
            cwd,
            _SW_SHOWNORMAL,
        )
    )
    return result > _SHELL_EXECUTE_SUCCESS_MIN


def _launch_updated_exe_subprocess(exe_path: Path, *, params: str, cwd: str) -> None:
    args = [str(exe_path)]
    if params:
        args.append(params)
    subprocess.Popen(
        args,
        cwd=cwd,
        close_fds=True,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def launch_updated_exe(exe_path: Path, *, show_changelog: bool, version: str) -> None:
    resolved = exe_path.resolve()
    params = f"--show-changelog={version}" if show_changelog else ""
    cwd = str(resolved.parent)
    if sys.platform == "win32":
        if _launch_updated_exe_windows(resolved, params=params, cwd=cwd):
            return
        _launch_updated_exe_subprocess(resolved, params=params, cwd=cwd)
        return
    args = [str(resolved)]
    if params:
        args.append(params)
    subprocess.Popen(args, close_fds=True, cwd=cwd)


def exit_for_update_restart() -> None:
    if sys.platform == "win32":
        time.sleep(2.0)
        os._exit(0)
    sys.exit(0)
