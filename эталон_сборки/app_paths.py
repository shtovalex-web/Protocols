# -*- coding: utf-8 -*-
"""Каталоги приложения: обычный запуск и сборка PyInstaller (.exe)."""

from __future__ import annotations

import sys
from pathlib import Path

_BUNDLE_MARKERS = (
    "default_protocol.docx",
    "default_protocol.odt",
)
# В исходниках шаблоны и справка — в bundle/; в поставке .exe — в подпапке data/ (не в корне с exe).
_BUNDLE_SUBDIR = "bundle"
_RESOURCE_DATA_SUBDIR = "data"
# Текстовый журнал ошибок UI и необработанных исключений (рядом с protocols.db).
ERROR_LOG_FILENAME = "protocol_errors_journal.txt"


def application_exe_dir() -> Path:
    """Каталог с исполняемым файлом (.exe) или каталог проекта при запуске из исходников."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def application_user_dir() -> Path:
    """
    Данные пользователя: protocols.db, Excel, Protokol/, Mintrud/ и t.д.

    Каталог с программой: рядом с .exe (сборка) или корень проекта (исходники).
    """
    return application_exe_dir()


def application_error_log_path() -> Path:
    """Путь к файлу журнала ошибок (UTF-8, дозапись)."""
    return application_user_dir() / ERROR_LOG_FILENAME


def _dir_has_bundle_marker(folder: Path) -> bool:
    return any((folder / name).is_file() for name in _BUNDLE_MARKERS)


def _first_dir_with_marker(candidates: list[Path]) -> Path | None:
    for folder in candidates:
        if folder.is_dir() and _dir_has_bundle_marker(folder):
            return folder
    return None


def application_bundle_dir() -> Path:
    """
    Встроенные шаблоны, образцы Excel, XSD Минтруда, справка (FAQ).

    Сборка .exe: сначала «data/» рядом с exe, затем файлы в корне (старые поставки).
    Исходники: bundle/, затем data/, затем корень проекта.
    """
    if not getattr(sys, "frozen", False):
        root = Path(__file__).resolve().parent
        found = _first_dir_with_marker(
            [
                root / _BUNDLE_SUBDIR,
                root / _RESOURCE_DATA_SUBDIR,
                root,
            ]
        )
        if found is not None:
            return found
        nested = root / _BUNDLE_SUBDIR
        return nested if nested.is_dir() else root

    exe_dir = application_exe_dir()
    candidates: list[Path] = [
        exe_dir / _RESOURCE_DATA_SUBDIR,
        exe_dir,
        exe_dir / "_internal" / _RESOURCE_DATA_SUBDIR,
        exe_dir / "_internal",
    ]
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        mp = Path(meipass)
        candidates.extend([mp / _RESOURCE_DATA_SUBDIR, mp])
    found = _first_dir_with_marker(candidates)
    if found is not None:
        return found
    return exe_dir / _RESOURCE_DATA_SUBDIR


def application_resource_data_subdir_name() -> str:
    """Имя подпапки комплекта шаблонов рядом с .exe (для сборки и документации)."""
    return _RESOURCE_DATA_SUBDIR
