# -*- coding: utf-8 -*-
"""Выгрузка комплекта файлов для восстановления рабочей среды."""

from __future__ import annotations

import json
import shutil
from pathlib import Path

from app_paths import (
    application_bundle_dir,
    application_exe_dir,
    application_resource_data_subdir_name,
)
from employees_io import (
    EMPLOYEES_EXCEL_FILENAME,
    PROGRAMS_EXCEL_FILENAME,
    write_template_data_base_workbook,
    write_template_programs_workbook,
)
from protocol_db import init_protocols_db_file
from protocol_docx import PROTOCOL_TEMPLATE_FILENAME
from protocol_paths import DATABASE_FILENAME, LAST_PROTOCOL_NO_STATE_FILENAME

RECOVERY_TEMPLATE_COPY_FILENAMES: tuple[str, ...] = (
    PROTOCOL_TEMPLATE_FILENAME,
    "default_protocol_tehnicheskiy.docx",
    "FAQ.txt",
    "FAQ.md",
    "icon.ico",
    "Шаблон_Минтруд_XSD_УМН.xlsx",
    "!! Шаблон_Минтруд_XSD_УМН _ общ+.xlsx",
    "Шаблон_Минтруд_XSD_УМН _ общ+.xlsx",
    "ПОДРОБНАЯ_ИНСТРУКЦИЯ_для_пользователя.docx",
    "ИНСТРУКЦИЯ_оформление_протоколов_Минтруд.docx",
)
RECOVERY_README_FILENAME = "ВОССТАНОВЛЕНИЕ_ДАННЫХ.txt"


def _recovery_bundle_readme_text() -> str:
    return f"""Шаблоны для восстановления работы программы
{'=' * 50}

Скопируйте нужные файлы, заменив утерянные или повреждённые. Перед заменой закройте программу.

При запуске из исходников — в каталог с main.py.
При запуске ProtocolOOT.exe:
  • шаблоны Word, XSD Минтруда, справка — в подпапку «{application_resource_data_subdir_name()}» рядом с .exe;
  • рабочие файлы ({DATABASE_FILENAME}, {EMPLOYEES_EXCEL_FILENAME}, {PROGRAMS_EXCEL_FILENAME},
    {LAST_PROTOCOL_NO_STATE_FILENAME}) — в корень (рядом с .exe), не в «{application_resource_data_subdir_name()}».

Файлы в этой выгрузке:
  • {DATABASE_FILENAME} — пустая база: журнал протоколов, настройки приказа/комиссии, реквизиты
    Минтруда, кэш листов Excel. После копирования заново укажите пути к файлам Excel в настройках,
    при необходимости восстановите приказ и комиссию.
  • {EMPLOYEES_EXCEL_FILENAME} — сотрудники и комиссия (без листов программ).
  • {PROGRAMS_EXCEL_FILENAME} — справочник программ (листы B, V_PROF, PP, SIZ, V); при отсутствии
    программы читаются из {EMPLOYEES_EXCEL_FILENAME}.
  • {LAST_PROTOCOL_NO_STATE_FILENAME} — сброшенный последний номер протокола (можно не копировать,
    если хотите сохранить текущий номер).
  • {PROTOCOL_TEMPLATE_FILENAME}, default_protocol_tehnicheskiy.docx — бланки протокола.
  • FAQ.txt (или FAQ.md) — справка в меню «Справка».
  • Шаблон_Минтруд*.xlsx — при наличии в выгрузке; официальный шаблон можно взять с портала Минтруда.
  • icon.ico — дополнительный значок Windows (необязательно; основной — вшитый в программу).

Файл программ можно держать отдельно ({PROGRAMS_EXCEL_FILENAME}) или в одном файле с сотрудниками.

Подробности — меню «Справка» и инструкции .docx в папке «{application_resource_data_subdir_name()}».
"""


def export_recovery_templates_to_folder(dest: Path) -> tuple[list[str], list[str]]:
    """
    Пишет в dest пустой protocols.db, шаблон Data_base.xlsx, last_protocol_no.json,
    копирует файлы комплекта (шаблон .docx, FAQ, XSD Минтруда и т.д.) и текстовую инструкцию.
    Возвращает (список имён созданных/скопированных файлов, список имён отсутствующих в комплекте).
    """
    dest = Path(dest).resolve()
    dest.mkdir(parents=True, exist_ok=True)
    data_subdir = application_resource_data_subdir_name()
    kit_dir = dest / data_subdir
    kit_dir.mkdir(parents=True, exist_ok=True)
    done: list[str] = []
    missing: list[str] = []
    root = application_exe_dir()
    bundle = application_bundle_dir()

    init_protocols_db_file(dest / DATABASE_FILENAME)
    done.append(DATABASE_FILENAME)

    write_template_data_base_workbook(dest / EMPLOYEES_EXCEL_FILENAME)
    done.append(EMPLOYEES_EXCEL_FILENAME)

    write_template_programs_workbook(dest / PROGRAMS_EXCEL_FILENAME)
    done.append(PROGRAMS_EXCEL_FILENAME)

    lp = dest / LAST_PROTOCOL_NO_STATE_FILENAME
    lp.write_text(
        json.dumps({"last_protocol_no": ""}, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    done.append(LAST_PROTOCOL_NO_STATE_FILENAME)

    for name in RECOVERY_TEMPLATE_COPY_FILENAMES:
        src = bundle / name
        if not src.is_file():
            src = root / data_subdir / name
        if not src.is_file():
            src = root / name
        if not src.is_file() and name == "FAQ.txt":
            faq_md = bundle / "FAQ.md"
            if not faq_md.is_file():
                faq_md = root / data_subdir / "FAQ.md"
            if not faq_md.is_file():
                faq_md = root / "FAQ.md"
            if faq_md.is_file():
                src = faq_md
        if not src.is_file():
            missing.append(name)
            continue
        out_name = "FAQ.txt" if src.name == "FAQ.md" and name == "FAQ.txt" else name
        shutil.copy2(src, kit_dir / out_name)
        done.append(f"{data_subdir}/{out_name}")

    (dest / RECOVERY_README_FILENAME).write_text(
        _recovery_bundle_readme_text(),
        encoding="utf-8",
    )
    done.append(RECOVERY_README_FILENAME)
    return done, missing
