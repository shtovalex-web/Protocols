# -*- coding: utf-8 -*-
"""
Архив исходников для проверки ИБ: код, bundle, сборка exe, без пользовательских данных.

    py -3 tools/pack_ib_review.py

Создаёт ib_review_YYYYMMDD_HHMMSS.zip в корне проекта.
"""

from __future__ import annotations

import os
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BUNDLE = ROOT / "bundle"
NEXT = ROOT / "ProtocolOHT_next"
TOOLS = ROOT / "tools"

ROOT_PY = (
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
    "v_prof_combinations.py",
    "russian_genitive.py",
)

ROOT_MISC = (
    "README.md",
    "requirements.txt",
    "requirements.lock",
    "requirements-build.txt",
    "requirements-dev.txt",
    "ruff.toml",
    ".gitignore",
    "build_windows_exe.py",
    "build_windows_exe.bat",
    "verify.bat",
    "update_etalon.bat",
    "generate_oformlenie_instruction_docx.bat",
    "generate_podrobnaya_instruction_docx.bat",
    "ProtocolOOT.spec",
)

TOOLS_PY = (
    "verify_project.py",
    "update_etalon.py",
    "instruction_md_to_docx.py",
    "generate_oformlenie_instruction_docx.py",
    "generate_podrobnaya_instruction_docx.py",
    "_gen_embedded_png.py",
    "repair_technical_template_docx.py",
    "repair_protocol_templates_docx.py",
    "protect_protocol_templates.py",
    "patch_protocol_template_markers.py",
    "remove_logo_background.py",
)

IB_README = """Пакет исходников ProtocolOOT для проверки информационной безопасности
================================================================================

Состав: только файлы, необходимые для запуска из исходников и сборки папки с программой.

НЕ включено (персональные / рабочие данные и артефакты):
  • protocols.db, last_protocol_no.json, protocol_errors_journal.txt
  • пользовательские Excel (рабочие Data_base.xlsx из корня)
  • папки Protokol/, Mintrud/, local/, эталон_сборки/
  • собранный exe, кэши (__pycache__, .ruff_cache), .git, .cursor

Включено:
  • исходный код Python (корень + ProtocolOHT_next/)
  • bundle/ — шаблоны Word, XSD Минтруда, справка, образец Programs_base.xlsx
  • Data_base.xlsx и Programs_base.xlsx в корне архива — пустые шаблоны (без сотрудников)
  • tools/ — проверка проекта, сборка эталона, генерация инструкций
  • build_windows_exe.py — сборка ProtocolOOT.exe + папка data/

Запуск из исходников:
  py -3 -m pip install -r requirements.txt
  py -3 main.py

Сборка exe (Windows):
  py -3 -m pip install -r requirements-build.txt
  py -3 build_windows_exe.py

Проверка перед сборкой:
  py -3 tools/verify_project.py --no-launch
"""


def _copy_file(src: Path, dest: Path) -> bool:
    if not src.is_file():
        return False
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dest)
    return True


def _zip_dir(folder: Path, zip_path: Path) -> None:
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _dirs, files in os.walk(folder):
            for name in files:
                full = Path(root) / name
                arc = full.relative_to(folder).as_posix()
                zf.write(full, arc)


def main() -> int:
    os.chdir(ROOT)
    sys.path.insert(0, str(ROOT))

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    staging = ROOT / f"_ib_staging_{stamp}"
    zip_path = ROOT / f"ib_review_{stamp}.zip"

    if staging.exists():
        shutil.rmtree(staging)
    staging.mkdir(parents=True)

    copied: list[str] = []
    missing: list[str] = []

    for name in ROOT_PY + ROOT_MISC:
        if _copy_file(ROOT / name, staging / name):
            copied.append(name)
        elif name in ROOT_PY or name in ("README.md", "requirements.txt", "build_windows_exe.py"):
            missing.append(name)

    if BUNDLE.is_dir():
        for src in BUNDLE.rglob("*"):
            if src.is_file():
                rel = src.relative_to(BUNDLE)
                dest = staging / "bundle" / rel
                _copy_file(src, dest)
                copied.append(f"bundle/{rel.as_posix()}")

    if NEXT.is_dir():
        for src in NEXT.glob("*.py"):
            rel = src.name
            if _copy_file(src, staging / "ProtocolOHT_next" / rel):
                copied.append(f"ProtocolOHT_next/{rel}")

    for name in TOOLS_PY:
        if _copy_file(TOOLS / name, staging / "tools" / name):
            copied.append(f"tools/{name}")

    try:
        from employees_io import (
            write_template_data_base_workbook,
            write_template_programs_workbook,
        )

        write_template_data_base_workbook(staging / "Data_base.xlsx")
        copied.append("Data_base.xlsx (шаблон)")
        prog_tpl = staging / "Programs_base.xlsx"
        if (BUNDLE / "Programs_base.xlsx").is_file():
            shutil.copy2(BUNDLE / "Programs_base.xlsx", prog_tpl)
            copied.append("Programs_base.xlsx (из bundle)")
        else:
            write_template_programs_workbook(prog_tpl)
            copied.append("Programs_base.xlsx (шаблон)")
    except Exception as e:
        print(f"Предупреждение: не удалось создать шаблоны Excel: {e}", file=sys.stderr)

    (staging / "ИБ_ПАМЯТКА.txt").write_text(IB_README, encoding="utf-8")
    copied.append("ИБ_ПАМЯТКА.txt")

    if missing:
        print("Не найдено (обязательное):", ", ".join(missing), file=sys.stderr)

    if zip_path.is_file():
        zip_path.unlink()
    _zip_dir(staging, zip_path)
    shutil.rmtree(staging)

    size_mb = zip_path.stat().st_size / (1024 * 1024)
    print(f"Готово: {zip_path}")
    print(f"Файлов в пакете: {len(copied)}, размер: {size_mb:.2f} МБ")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
