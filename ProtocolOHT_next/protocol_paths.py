# -*- coding: utf-8 -*-
"""Пути к данным приложения: SQLite, состояние номера протокола."""

from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

from app_paths import (
    ERROR_LOG_FILENAME,
    application_bundle_dir,
    application_exe_dir,
    application_user_dir,
)
from bundle_integration import bundle_spreadsheet_path, resolve_openpyxl_workbook_path
from employees_io import EMPLOYEES_EXCEL_FILENAME, PROGRAMS_EXCEL_FILENAME

DATABASE_FILENAME = "protocols.db"
LAST_PROTOCOL_NO_STATE_FILENAME = "last_protocol_no.json"
PROTOCOL_OUTPUT_SUBDIR = "Protokol"
MINTRUD_EXPORT_SUBDIR = "Mintrud"
# Старая схема: рабочие файлы в подпапке «Данные»; при откате поднимаем в каталог с .exe.
_LEGACY_DATA_SUBDIR = "Данные"


def database_path() -> Path:
    return application_user_dir() / DATABASE_FILENAME


def employees_excel_default_path() -> Path:
    """Рабочий Excel в каталоге данных; при отсутствии — образец из bundle (.xlsx или .ods)."""
    u = application_user_dir() / EMPLOYEES_EXCEL_FILENAME
    if u.is_file():
        return resolve_openpyxl_workbook_path(u)
    b = bundle_spreadsheet_path("Data_base")
    if b is not None:
        return b
    legacy = application_bundle_dir() / EMPLOYEES_EXCEL_FILENAME
    if legacy.is_file():
        return resolve_openpyxl_workbook_path(legacy)
    return resolve_openpyxl_workbook_path(u)


def programs_excel_default_path() -> Path:
    """
    Справочник программ (B, V_PROF, PP, SIZ, V). Если Programs_base нет рядом с программой —
    используется тот же файл, что и для сотрудников (объединённый Data_base).
    """
    u = application_user_dir() / PROGRAMS_EXCEL_FILENAME
    if u.is_file():
        return resolve_openpyxl_workbook_path(u)
    b = bundle_spreadsheet_path("Programs_base")
    if b is not None:
        return b
    legacy = application_bundle_dir() / PROGRAMS_EXCEL_FILENAME
    if legacy.is_file():
        return resolve_openpyxl_workbook_path(legacy)
    return employees_excel_default_path()


def protocols_output_dir() -> Path:
    """Папка для сохранения протоколов PDF/DOCX; создаётся при отсутствии."""
    p = application_user_dir() / PROTOCOL_OUTPUT_SUBDIR
    p.mkdir(parents=True, exist_ok=True)
    return p


def mintrud_export_output_dir() -> Path:
    """Папка для выгрузки Excel-шаблона Минтруда; создаётся при отсутствии."""
    p = application_user_dir() / MINTRUD_EXPORT_SUBDIR
    p.mkdir(parents=True, exist_ok=True)
    return p


def last_protocol_no_state_path() -> Path:
    return application_user_dir() / LAST_PROTOCOL_NO_STATE_FILENAME


def load_last_protocol_no() -> str:
    p = last_protocol_no_state_path()
    if not p.is_file():
        return ""
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError, TypeError):
        return ""
    if not isinstance(data, dict):
        return ""
    v = data.get("last_protocol_no", "")
    return v.strip() if isinstance(v, str) else ""


def save_last_protocol_no(value: str) -> None:
    p = last_protocol_no_state_path()
    try:
        p.write_text(
            json.dumps({"last_protocol_no": value}, ensure_ascii=False),
            encoding="utf-8",
        )
    except OSError:
        pass


def migrate_legacy_from_data_subfolder_to_exe() -> None:
    """
    Однократно поднимает protocols.db, Excel, журнал, Protokol/ и Mintrud/ из папки «Данные»
    в каталог с .exe, если в корне каталога exe базы ещё нет (откат со старой схемы хранения).
    """
    if not getattr(sys, "frozen", False):
        return
    exe_dir = application_exe_dir()
    legacy = exe_dir / _LEGACY_DATA_SUBDIR
    if not legacy.is_dir():
        return
    if (exe_dir / DATABASE_FILENAME).is_file():
        return

    def move_up(name: str) -> None:
        src = legacy / name
        dst = exe_dir / name
        if not src.exists() or dst.exists():
            return
        shutil.move(str(src), str(dst))

    move_up(DATABASE_FILENAME)
    move_up(LAST_PROTOCOL_NO_STATE_FILENAME)
    move_up(ERROR_LOG_FILENAME)
    move_up(EMPLOYEES_EXCEL_FILENAME)
    move_up(PROGRAMS_EXCEL_FILENAME)
    move_up(PROTOCOL_OUTPUT_SUBDIR)
    move_up(MINTRUD_EXPORT_SUBDIR)
    try:
        if legacy.is_dir() and not any(legacy.iterdir()):
            legacy.rmdir()
    except OSError:
        pass


def ensure_frozen_default_workbooks() -> None:
    """
    В сборке .exe создаёт пустой Data_base рядом с exe, если файла нет.

    Programs_base.xlsx здесь намеренно не создаётся: иначе появится пустой файл и программы
    перестанут подхватываться из объединённого Data_base (см. programs_excel_default_path).
    """
    if not getattr(sys, "frozen", False):
        return
    from employees_io import write_template_data_base_workbook

    emp = application_user_dir() / EMPLOYEES_EXCEL_FILENAME
    if not emp.is_file():
        write_template_data_base_workbook(emp)
