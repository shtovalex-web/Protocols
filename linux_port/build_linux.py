# -*- coding: utf-8 -*-
"""
Сборка приложения в один исполняемый файл (PyInstaller --onefile, Linux).

Запускать на Linux после prepare.py:
    python3 linux_port/build_linux.py
    python3 build_linux.py          # из linux_port/
    ./build_linux.sh

Результат: linux_port/ProtocolOHT_linux_dist/ — ProtocolOOT + data/
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

LINUX_PORT = Path(__file__).resolve().parent
APP = LINUX_PORT / "app"
BUNDLE_DIR = APP / "bundle"

sys.path.insert(0, str(APP))
from app_paths import application_resource_data_subdir_name  # noqa: E402

DATA_SUBDIR = application_resource_data_subdir_name()
NEXT = APP / "ProtocolOHT_next"
DEFAULT_OUT_DIR = LINUX_PORT / "ProtocolOHT_linux_dist"
EXE_NAME = "ProtocolOOT"
WORK = LINUX_PORT / "_pyinstaller_build_linux"

_PYI_HIDDEN = [
    "app_paths",
    "program_keys",
    "clipboard_ui",
    "commission_admin",
    "employees_io",
    "excel_data_cache",
    "docx_template_protection",
    "programs_v_prof",
    "v_prof_combinations",
    "faq_viewer",
    "mintrud_export",
    "mintrud_trained_registry",
    "v_program_registry_match",
    "russian_genitive",
    "fpdf",
    "protocol_db",
    "protocol_errors",
    "protocol_paths",
    "protocol_journal",
    "protocol_docx",
    "protocol_output",
    "protocol_recovery",
    "protocol_app_info",
    "protocol_ui",
    "protocol_embedded_assets",
]

_EMPLOYEES_XLSX = "Data_base.xlsx"
_PROGRAMS_XLSX = "Programs_base.xlsx"

BUNDLE_FILES = [
    "default_protocol.docx",
    "default_protocol_tehnicheskiy.docx",
    "ПОДРОБНАЯ_ИНСТРУКЦИЯ_для_пользователя.docx",
    "ИНСТРУКЦИЯ_оформление_протоколов_Минтруд.docx",
    _EMPLOYEES_XLSX,
    _PROGRAMS_XLSX,
    "icon.ico",
    "Шаблон_Минтруд_XSD_УМН.xlsx",
    "!! Шаблон_Минтруд_XSD_УМН _ общ+.xlsx",
    "Шаблон_Минтруд_XSD_УМН _ общ+.xlsx",
]
BUNDLE_EXE_ROOT_XLSX = (_EMPLOYEES_XLSX, _PROGRAMS_XLSX)

DIST_README = """Папка готовой Linux-сборки (onefile + комплект в data/)

• ProtocolOOT — программа (один исполняемый файл).
• Подпапка data/ — шаблоны Word, образцы Excel, XSD Минтруда, справка FAQ.txt, инструкции (.docx).

Для PDF с оформлением DOCX на целевом ПК нужен LibreOffice:
  sudo apt install libreoffice-writer fonts-dejavu-core

Упрощённый PDF из предпросмотра использует системные TTF (DejaVu, Liberation и т.п.).

Переносите на другой ПК всю папку: ProtocolOOT + data/ целиком.
"""


def _bundle_src(name: str) -> Path:
    p = BUNDLE_DIR / name
    if p.is_file():
        return p
    return APP / name


def _run_verify() -> int:
    script = LINUX_PORT / "verify_linux.py"
    if not script.is_file():
        return 0
    print("Проверка Linux-копии (verify_linux.py --no-launch)...")
    return subprocess.run(
        [sys.executable, str(script), "--no-launch"],
        cwd=str(LINUX_PORT),
    ).returncode


def _run_ruff() -> int:
    rc = subprocess.run(
        [sys.executable, "-m", "ruff", "check", str(APP)],
        cwd=str(LINUX_PORT),
    ).returncode
    if rc != 0:
        print("Сборка остановлена: исправьте замечания ruff.", file=sys.stderr)
    return rc


def main() -> int:
    if not (APP / "main.py").is_file():
        print(
            f"Нет {APP}/main.py — сначала выполните: python linux_port/prepare.py",
            file=sys.stderr,
        )
        return 1

    os.chdir(APP)
    sys.path.insert(0, str(NEXT))
    sys.path.insert(1, str(APP))

    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print(
            "Установите PyInstaller: python3 -m pip install -r linux_port/requirements-build.txt",
            file=sys.stderr,
        )
        return 1

    if _run_verify() != 0:
        return 1
    if _run_ruff() != 0:
        return 1

    if len(sys.argv) > 1:
        OUT_DIR = Path(" ".join(sys.argv[1:]).strip().strip('"')).expanduser().resolve()
    else:
        OUT_DIR = DEFAULT_OUT_DIR
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    args: list[str] = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        f"--name={EXE_NAME}",
        f"--distpath={OUT_DIR}",
        f"--workpath={WORK}",
        f"--specpath={LINUX_PORT}",
        f"--paths={NEXT}",
        f"--paths={APP}",
    ]
    ico = _bundle_src("icon.ico")
    if ico.is_file():
        args.append(f"--icon={ico}")

    for mod in _PYI_HIDDEN:
        args.append(f"--hidden-import={mod}")
    args.extend(
        [
            "--hidden-import=docx",
            "--hidden-import=docx.oxml",
            "--collect-submodules=openpyxl",
            "--collect-submodules=pymorphy2",
        ]
    )
    try:
        import pymorphy2_dicts_ru  # noqa: F401
    except ImportError:
        pass
    else:
        args.append("--collect-data=pymorphy2_dicts_ru")

    args.append(str(APP / "main.py"))

    print("PyInstaller: --onefile, выход:", OUT_DIR)
    r = subprocess.run(args, cwd=str(APP))
    if r.returncode != 0:
        return r.returncode

    binary = OUT_DIR / EXE_NAME
    if not binary.is_file():
        print("Ошибка: не найден", binary, file=sys.stderr)
        return 1
    binary.chmod(binary.stat().st_mode | 0o111)

    data_dir = OUT_DIR / DATA_SUBDIR
    data_dir.mkdir(parents=True, exist_ok=True)
    fonts_dir = data_dir / "fonts"
    fonts_dir.mkdir(parents=True, exist_ok=True)

    copied = 0
    for name in BUNDLE_FILES:
        src = _bundle_src(name)
        if src.is_file():
            shutil.copy2(src, data_dir / name)
            copied += 1
            if name in BUNDLE_EXE_ROOT_XLSX:
                shutil.copy2(src, OUT_DIR / name)

    faq_src = _bundle_src("FAQ.md")
    if faq_src.is_file():
        shutil.copy2(faq_src, data_dir / "FAQ.txt")
        copied += 1

    (OUT_DIR / "ИНСТРУКЦИЯ_папки_сборки.txt").write_text(DIST_README, encoding="utf-8")

    print()
    print("Сборка завершена.")
    print(f"  {binary}")
    print(f"  Комплект в {data_dir.name}/: {copied} файл(ов)")
    print()
    print(f"Переносите на другие ПК всю папку:\n  {OUT_DIR}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
