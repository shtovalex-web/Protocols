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
from windows_app_bundle import (
    APP_UPDATE_STAGING_DIR,
    DEFAULT_EXE_NAME,
    INTERNAL_DIR_NAME,
    WINDOWS_APP_ZIP_NAME,
    app_update_staging_dir,
    staged_app_zip_path,
)

# ShellExecuteW: > 32 — успех, иначе код ошибки Windows.
_SHELL_EXECUTE_SUCCESS_MIN = 32
_SW_SHOWNORMAL = 1
_RESTART_CMD_NAME = "ProtocolOOT_restart.cmd"


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


def stage_windows_app_bundle(
    install_dir: Path,
    source: Path,
    *,
    expected_sha256: str,
    expected_size: int,
) -> Path:
    """Скачать zip, распаковать в .app_update_staging/ (применение — при следующем запуске exe)."""
    root = install_dir.resolve()
    staged_zip = staged_app_zip_path(root)
    stage_payload_copy(
        source,
        staged_zip,
        expected_sha256=expected_sha256,
        expected_size=expected_size,
    )
    staging = app_update_staging_dir(root)
    if staging.exists():
        shutil.rmtree(staging, ignore_errors=True)
    staging.mkdir(parents=True, exist_ok=True)
    shutil.unpack_archive(staged_zip, staging)
    exe = staging / DEFAULT_EXE_NAME
    internal = staging / INTERNAL_DIR_NAME
    if not exe.is_file() or not internal.is_dir():
        shutil.rmtree(staging, ignore_errors=True)
        staged_zip.unlink(missing_ok=True)
        msg = f"Invalid bundle archive: {source.name}"
        raise UpdateInstallerError(msg)
    return staging


def cleanup_app_update_staging(install_dir: Path) -> None:
    root = install_dir.resolve()
    staging = app_update_staging_dir(root)
    if staging.is_dir():
        shutil.rmtree(staging, ignore_errors=True)
    staged_zip = staged_app_zip_path(root)
    if staged_zip.is_file():
        staged_zip.unlink(missing_ok=True)


def apply_pending_app_staging(install_dir: Path) -> bool:
    """Применить .app_update_staging/ после ручного перезапуска (zip-manifest)."""
    root = install_dir.resolve()
    staging = app_update_staging_dir(root)
    staged_exe = staging / DEFAULT_EXE_NAME
    if not staged_exe.is_file():
        return False
    for item in staging.iterdir():
        dest = root / item.name
        if item.is_dir():
            if dest.exists():
                shutil.rmtree(dest, ignore_errors=True)
            shutil.copytree(item, dest)
        else:
            shutil.copy2(item, dest)
    cleanup_app_update_staging(root)
    return True


def restart_cmd_path(exe_path: Path) -> Path:
    return exe_path.resolve().parent / _RESTART_CMD_NAME


def _write_restart_cmd(exe_path: Path, *, parent_pid: int) -> Path:
    """Локальный .cmd рядом с .exe: только ASCII, путь через %~dp0 (кириллица и «!»)."""
    resolved = exe_path.resolve()
    if not resolved.is_file():
        msg = f"Executable not found: {resolved}"
        raise UpdateInstallerError(msg)
    if parent_pid <= 0:
        msg = f"Invalid parent PID: {parent_pid}"
        raise UpdateInstallerError(msg)
    cmd_path = restart_cmd_path(resolved)
    pid = int(parent_pid)
    # Ждём завершения onefile-родителя — иначе второй _MEI ломает python3xx.dll.
    content = (
        "@echo off\r\n"
        "setlocal DisableDelayedExpansion\r\n"
        "set TCL_LIBRARY=\r\n"
        "set TK_LIBRARY=\r\n"
        "set TCLLIBPATH=\r\n"
        "set TKPATH=\r\n"
        'cd /d "%~dp0"\r\n'
        "set /a _wait=0\r\n"
        ":wait_parent\r\n"
        "set /a _wait+=1\r\n"
        "if _wait gtr 90 goto start_app\r\n"
        f'tasklist /FI "PID eq {pid}" 2>nul | find "{pid}" >nul\r\n'
        "if not errorlevel 1 (\r\n"
        "  ping 127.0.0.1 -n 2 >nul\r\n"
        "  goto wait_parent\r\n"
        ")\r\n"
        ":start_app\r\n"
        f'if exist "%~dp0{APP_UPDATE_STAGING_DIR}\\{DEFAULT_EXE_NAME}" (\r\n'
        f'  robocopy "%~dp0{APP_UPDATE_STAGING_DIR}" "%~dp0" /E /IS /IT /NFL /NDL /NJH /NJS /NC /NS\r\n'
        f'  if exist "%~dp0{APP_UPDATE_STAGING_DIR}" rmdir /s /q "%~dp0{APP_UPDATE_STAGING_DIR}"\r\n'
        ")\r\n"
        f'if exist "%~dp0{WINDOWS_APP_ZIP_NAME}.new" del /q "%~dp0{WINDOWS_APP_ZIP_NAME}.new"\r\n'
        f'start "" /D "%~dp0" "%~dp0{resolved.name}"\r\n'
        'del "%~f0"\r\n'
    )
    cmd_path.write_text(content, encoding="ascii")
    return cmd_path


def cleanup_restart_cmd(exe_path: Path) -> bool:
    """Удалить оставшийся ProtocolOOT_restart.cmd после перезапуска."""
    cmd_path = restart_cmd_path(exe_path)
    if not cmd_path.is_file():
        return True
    return _unlink_with_retries(cmd_path)


def _shell_execute(path: Path, *, cwd: str) -> bool:
    import ctypes

    result = int(
        ctypes.windll.shell32.ShellExecuteW(
            None,
            "open",
            str(path),
            None,
            cwd,
            _SW_SHOWNORMAL,
        )
    )
    return result > _SHELL_EXECUTE_SUCCESS_MIN


def _launch_updated_exe_cmd_helper(exe_path: Path) -> bool:
    """Запуск через локальный .cmd (%~dp0) — без кириллицы в теле сценария."""
    try:
        cmd_path = _write_restart_cmd(exe_path, parent_pid=os.getpid())
    except (OSError, UpdateInstallerError):
        return False
    return _shell_execute(cmd_path, cwd=str(cmd_path.parent))


def _launch_updated_exe_windows(exe_path: Path, *, cwd: str) -> bool:
    """ShellExecuteW напрямую на .exe; True при успехе."""
    if not exe_path.is_file():
        return False
    return _shell_execute(exe_path, cwd=cwd)


def _launch_updated_exe_subprocess(exe_path: Path, *, cwd: str) -> None:
    subprocess.Popen(
        [str(exe_path)],
        cwd=cwd,
        close_fds=True,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def launch_updated_exe(exe_path: Path) -> None:
    """Запустить обновлённый .exe без аргументов (changelog — через .pending_changelog.json)."""
    resolved = exe_path.resolve()
    cwd = str(resolved.parent)
    if sys.platform == "win32":
        if _launch_updated_exe_cmd_helper(resolved):
            return
        if _launch_updated_exe_windows(resolved, cwd=cwd):
            return
        _launch_updated_exe_subprocess(resolved, cwd=cwd)
        return
    subprocess.Popen([str(resolved)], close_fds=True, cwd=cwd)


def exit_after_update() -> None:
    """Завершить процесс после установки обновления (без автозапуска нового exe)."""
    if sys.platform == "win32":
        os._exit(0)
    sys.exit(0)


def exit_for_update_restart() -> None:
    """Устаревшее имя; оставлено для совместимости."""
    exit_after_update()
