# -*- coding: utf-8 -*-
"""Проверка обновлений при старте (MVP: замена .exe через rename)."""

from __future__ import annotations

import os
import sys
from pathlib import Path

from tkinter import messagebox

from pending_changelog import write_pending_changelog
from update_info import data_dir_for_exe
from protocol_app_info import APP_VERSION
from update_config import ENV_FORCE_CHECK, load_update_config, resolve_update_share_root
from update_data_installer import apply_data_updates, mark_data_version_installed
from update_installer import (
    UpdateInstallerError,
    apply_pending_app_staging,
    cleanup_backup_exe,
    cleanup_restart_cmd,
    stage_payload_copy,
    stage_windows_app_bundle,
    staged_new_exe_path,
    swap_exe_via_rename,
)
from update_success import notify_update_success_and_exit
from windows_app_bundle import is_windows_app_bundle
from update_manifest import UpdateManifestError
from update_scan import resolve_latest_update


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def parse_changelog_version(argv: list[str]) -> str | None:
    for arg in argv[1:]:
        if arg.startswith("--show-changelog="):
            return arg.partition("=")[2] or None
    return None


def should_check_for_updates(*, force: bool = False) -> bool:
    if force:
        return True
    if os.environ.get(ENV_FORCE_CHECK) == "1":
        return True
    return is_frozen()


def current_exe_path() -> Path | None:
    if not is_frozen():
        return None
    return Path(sys.executable).resolve()


def app_version() -> str:
    exe = current_exe_path()
    if exe is not None:
        from update_info import installed_version_from_data

        data_version = installed_version_from_data(exe)
        if data_version:
            return data_version
    return (APP_VERSION or "").strip()


def _load_changelog_items(version: str) -> list[str]:
    try:
        config = load_update_config()
        if not config.enabled:
            return []
        resolved = resolve_latest_update(
            resolve_update_share_root(config.manifest_path),
            current_version=app_version(),
        )
        if resolved is not None and resolved.version == version:
            return list(resolved.manifest.changes_short)
    except (UpdateManifestError, OSError):
        return []
    return []


def _resolve_update(config) -> tuple[Path, object] | None:
    resolved = resolve_latest_update(
        resolve_update_share_root(config.manifest_path),
        current_version=app_version(),
    )
    if resolved is None:
        return None
    return resolved.anchor_manifest_path, resolved.manifest


def _ask_install_update(
    manifest_version: str,
    changes: list[str],
    *,
    mandatory: bool,
    parent=None,
) -> bool:
    lines = [f"Доступна новая версия {manifest_version}."]
    if changes:
        lines.append("")
        lines.extend(f"• {item}" for item in changes)
    lines.append("")
    lines.append("Будут обновлены программа и файлы в папке data/ (шаблоны, справка).")
    lines.append("Файлы в корне папки (базы, protocols.db) не изменяются.")
    lines.append("Перед установкой закройте Word и документы из data/.")
    lines.append("")
    lines.append("Установить обновление сейчас?")
    text = "\n".join(lines)
    if mandatory:
        return messagebox.askokcancel("Обновление программы", text, parent=parent)
    return messagebox.askyesno("Обновление программы", text, parent=parent)


def _perform_update(manifest_path: Path, manifest, exe_path: Path, *, parent=None) -> None:
    source = manifest.windows_payload_path(manifest_path)
    install_dir = exe_path.resolve().parent
    bundle_staged = is_windows_app_bundle(source)
    if bundle_staged:
        stage_windows_app_bundle(
            install_dir,
            source,
            expected_sha256=manifest.windows.sha256,
            expected_size=manifest.windows.size,
        )
    else:
        stage_payload_copy(
            source,
            staged_new_exe_path(exe_path),
            expected_sha256=manifest.windows.sha256,
            expected_size=manifest.windows.size,
        )
    apply_data_updates(manifest_path, manifest, exe_path)
    mark_data_version_installed(exe_path, manifest)
    if not bundle_staged:
        swap_exe_via_rename(exe_path)
    write_pending_changelog(
        data_dir_for_exe(exe_path),
        version=manifest.latest_version,
        changes=list(manifest.changes_short),
    )
    notify_update_success_and_exit(
        version=manifest.latest_version,
        exe_path=exe_path,
        parent=parent,
        bundle_staged=bundle_staged,
    )


def _run_update_check(*, force: bool, parent=None) -> bool:
    """True — продолжить работу; False — выход после установки обновления."""
    if not should_check_for_updates(force=force):
        return True

    config = load_update_config()
    if not config.enabled:
        if force and parent is not None:
            messagebox.showinfo(
                "Обновление",
                "Проверка обновлений отключена в update_config.json.",
                parent=parent,
            )
        return True

    share_root = resolve_update_share_root(config.manifest_path)
    if not share_root.is_dir():
        if force and parent is not None:
            messagebox.showinfo(
                "Обновление",
                f"Не удалось прочитать каталог обновлений:\n{share_root}",
                parent=parent,
            )
        return True

    try:
        resolved = _resolve_update(config)
    except (UpdateManifestError, OSError):
        resolved = None

    if resolved is None:
        if force and parent is not None:
            messagebox.showinfo(
                "Обновление",
                f"Установлена актуальная версия ({app_version()}).\n"
                f"Каталог: {share_root}",
                parent=parent,
            )
        return True

    manifest_path, manifest = resolved

    exe_path = current_exe_path()
    if exe_path is None:
        if force and parent is not None:
            messagebox.showinfo(
                "Обновление",
                "Установка обновления доступна только для собранной программы (.exe).",
                parent=parent,
            )
        return True

    if not _ask_install_update(
        manifest.latest_version,
        manifest.changes_short,
        mandatory=manifest.mandatory,
        parent=parent,
    ):
        return True

    try:
        _perform_update(manifest_path, manifest, exe_path, parent=parent)
    except (UpdateInstallerError, OSError) as error:
        messagebox.showerror(
            "Обновление",
            f"Не удалось установить обновление:\n{error}",
            parent=parent,
        )
        cleanup_backup_exe(exe_path)
        return True

    return False


def check_updates_interactive(parent=None) -> None:
    """Ручная проверка из меню «Справка»."""
    _run_update_check(force=True, parent=parent)


def prepare_startup_updates(argv: list[str]) -> bool:
    """False — процесс завершается после установки обновления при старте."""
    changelog_version = parse_changelog_version(argv)
    if changelog_version:
        exe_path = current_exe_path()
        if exe_path is not None:
            write_pending_changelog(
                data_dir_for_exe(exe_path),
                version=changelog_version,
                changes=_load_changelog_items(changelog_version),
            )

    exe_path = current_exe_path()
    if exe_path is not None:
        apply_pending_app_staging(exe_path.parent)
        cleanup_backup_exe(exe_path)
        cleanup_restart_cmd(exe_path)

    return _run_update_check(force=False)
