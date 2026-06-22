# -*- coding: utf-8 -*-
"""Обновить папку «эталон_сборки» из файлов репозитория.

Запуск:
    py -3 tools/update_etalon.py
    или двойной щелчок по update_etalon.bat (корень проекта)

Не копирует: protocols.db, last_protocol_no.json, __pycache__, .venv и пр.
Комплект шаблонов — в эталон/data/ (как у .exe); Data_base/Programs_base — ещё в корне эталона.
"""

from __future__ import annotations

import os
import shutil
import stat
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BUNDLE = ROOT / "bundle"
ETALON = ROOT / "эталон_сборки"
DATA_SUBDIR = "data"

PY_MODULES = (
    "main.py",
    "app_paths.py",
    "clipboard_ui.py",
    "commission_admin.py",
    "docx_template_protection.py",
    "employees_io.py",
    "excel_data_cache.py",
    "faq_viewer.py",
    "mintrud_export.py",
    "mintrud_trained_registry.py",
    "program_keys.py",
    "programs_v_prof.py",
    "v_program_registry_match.py",
    "russian_genitive.py",
)

# В data/ (как в поставке exe): без .md
DATA_KIT_FILES = (
    "default_protocol.docx",
    "default_protocol_tehnicheskiy.docx",
    "ПОДРОБНАЯ_ИНСТРУКЦИЯ_для_пользователя.docx",
    "ИНСТРУКЦИЯ_оформление_протоколов_Минтруд.docx",
    "Data_base.xlsx",
    "Programs_base.xlsx",
    "icon.ico",
    "Шаблон_Минтруд_XSD_УМН.xlsx",
    "!! Шаблон_Минтруд_XSD_УМН _ общ+.xlsx",
    "Шаблон_Минтруд_XSD_УМН _ общ+.xlsx",
)

# Дополнительно в корень эталона (для запуска из исходников и как у exe)
ETALON_ROOT_XLSX = ("Data_base.xlsx", "Programs_base.xlsx")

ASSETS_OPTIONAL_ROOT = (
    "README.md",
    "requirements.txt",
    "requirements.lock",
    "requirements-dev.txt",
)


def _bundle_src(name: str) -> Path:
    p = BUNDLE / name
    if p.is_file():
        return p
    return ROOT / name


ZAPUSK_BAT = """@echo off
chcp 65001 >nul
cd /d "%~dp0"
py -3 main.py
if errorlevel 1 pause
"""

USTANOVKA_BAT = """@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Установка пакетов из requirements.txt ...
py -3 -m pip install -r requirements.txt
echo.
echo Для PDF через Word: Microsoft Word и py -3 -m pip install pywin32
pause
"""

INSTRUKTSIYA = f"""Эталонная копия программы «формирование протоколов» (резервная сборка)

Состав: модули .py из корня, README.md, requirements*, каталог {DATA_SUBDIR}/ (шаблоны, FAQ.txt, XSD),
в корне — Data_base.xlsx и Programs_base.xlsx (как после сборки .exe).

В репозитории шаблоны лежат в bundle/; update_etalon копирует комплект в {DATA_SUBDIR}/.

На новом ПК:
1) Установите Python 3 с официального сайта (с компонентом tkinter).
2) Запустите установка_зависимостей.bat или вручную:
   py -3 -m pip install -r requirements.txt
3) Для PDF как в документе Word: установите Microsoft Word и выполните
   py -3 -m pip install pywin32
4) Запуск: запуск.bat или команда  py -3 main.py  из этой папки.

База protocols.db и файл last_protocol_no.json создаются при работе сами.
Папки Protokol и Mintrud создаются программой рядом с main.py.

Обновление эталона: update_etalon.bat или  py -3 tools/update_etalon.py

Источник правды — корень репозитория, ProtocolOHT_next и bundle/; эталон пересобирается скриптом.
"""


def _rmtree_robust(target: Path) -> None:
    """Удаление каталога на Windows: снятие read-only и несколько попыток при блокировке."""

    def _onexc(func, path, exc):
        if isinstance(exc, PermissionError):
            try:
                os.chmod(path, stat.S_IWRITE)
                func(path)
            except OSError as e2:
                raise exc from e2
        else:
            raise exc

    if not target.exists():
        return
    last: OSError | None = None
    for _ in range(5):
        try:
            if sys.version_info >= (3, 12):
                shutil.rmtree(target, onexc=_onexc)
            else:

                def _legacy(func, path, exc_info):
                    _onexc(func, path, exc_info[1])

                shutil.rmtree(target, onerror=_legacy)
            return
        except OSError as e:
            last = e
            time.sleep(0.35)
    if last:
        raise last


def main() -> int:
    os.chdir(ROOT)
    ETALON.mkdir(parents=True, exist_ok=True)
    data_dir = ETALON / DATA_SUBDIR

    copied: list[str] = []
    missing_required: list[str] = []
    missing_optional: list[str] = []

    for name in PY_MODULES:
        src = ROOT / name
        if src.is_file():
            shutil.copy2(src, ETALON / name)
            copied.append(name)
        else:
            missing_required.append(name)

    for name in ASSETS_OPTIONAL_ROOT:
        src = ROOT / name
        if src.is_file():
            shutil.copy2(src, ETALON / name)
            copied.append(name)
        else:
            missing_optional.append(name)

    if data_dir.exists():
        _rmtree_robust(data_dir)
    data_dir.mkdir(parents=True, exist_ok=True)

    for name in DATA_KIT_FILES:
        src = _bundle_src(name)
        if src.is_file():
            shutil.copy2(src, data_dir / name)
            copied.append(f"{DATA_SUBDIR}/{name}")
            if name in ETALON_ROOT_XLSX:
                shutil.copy2(src, ETALON / name)
                copied.append(name)
        else:
            missing_optional.append(f"{DATA_SUBDIR}/{name}")

    faq_src = _bundle_src("FAQ.md")
    if faq_src.is_file():
        shutil.copy2(faq_src, data_dir / "FAQ.txt")
        copied.append(f"{DATA_SUBDIR}/FAQ.txt")

    changelog_src = _bundle_src("ЖУРНАЛ_ДОРАБОТОК.md")
    if changelog_src.is_file():
        shutil.copy2(changelog_src, ETALON / "ЖУРНАЛ_ДОРАБОТОК.md")
        shutil.copy2(changelog_src, data_dir / "ЖУРНАЛ_ДОРАБОТОК.md")
        copied.append("ЖУРНАЛ_ДОРАБОТОК.md")
        copied.append(f"{DATA_SUBDIR}/ЖУРНАЛ_ДОРАБОТОК.md")

    next_src = ROOT / "ProtocolOHT_next"
    next_dst = ETALON / "ProtocolOHT_next"
    if next_src.is_dir():
        if next_dst.exists():
            _rmtree_robust(next_dst)
        shutil.copytree(
            next_src,
            next_dst,
            ignore=shutil.ignore_patterns("__pycache__", "*.pyc", "desktop.ini"),
        )
        copied.append("ProtocolOHT_next/")

    (ETALON / "запуск.bat").write_text(ZAPUSK_BAT, encoding="utf-8", newline="\r\n")
    (ETALON / "установка_зависимостей.bat").write_text(
        USTANOVKA_BAT, encoding="utf-8", newline="\r\n"
    )
    (ETALON / "ИНСТРУКЦИЯ.txt").write_text(INSTRUKTSIYA, encoding="utf-8", newline="\r\n")
    copied.extend(["запуск.bat", "установка_зависимостей.bat", "ИНСТРУКЦИЯ.txt"])

    junk = ETALON / "desktop.ini"
    if junk.is_file():
        try:
            junk.unlink()
        except OSError:
            pass

    print(f"Эталон: {ETALON}")
    print(f"Скопировано/записано: {len(copied)} элементов.")
    if missing_required:
        print("ОШИБКА: нет обязательных файлов:", ", ".join(missing_required))
        return 1
    if missing_optional:
        print("Нет в исходниках (не критично):", ", ".join(missing_optional))
    return 0


if __name__ == "__main__":
    sys.exit(main())
