# -*- coding: utf-8 -*-
"""
Сборка приложения в один .exe (PyInstaller --onefile, Windows).

Шаблоны и (при наличии) Data_base.xlsx / Programs_base.xlsx не вшиваются в exe:
после сборки копируются в подпапку data/ рядом с .exe (в корне — только exe и рабочие файлы).
Рабочие файлы (protocols.db, Protokol/ и т.д.) создаются рядом с .exe (см. app_paths.application_user_dir).

Результат: каталог с exe и файлами рядом (по умолчанию ProtocolOHT_onefile/ в корне проекта).

Запуск:
    py -3 -m pip install -r requirements-build.txt
    py -3 build_windows_exe.py
    (перед PyInstaller выполняется ruff check . по ruff.toml; только проверка: verify.bat → tools/verify_project.py)
    py -3 build_windows_exe.py "D:\\Проекты Курсор\\Программа протокола"
    или двойной щелчок по build_windows_exe.bat (рабочая папка — каталог со скриптом)

Без аргументов сначала открывается диалог выбора папки (tkinter); «Отмена» — выход без сборки.
Необязательный аргумент — путь к папке вывода (перетаскивание на .bat, создание при необходимости).

Если PyInstaller не ставится на очень новый Python (например 3.14), соберите на ПК с Python 3.11–3.12.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
BUNDLE_DIR = ROOT / "bundle"

from app_paths import application_resource_data_subdir_name

DATA_SUBDIR = application_resource_data_subdir_name()


def _bundle_src(name: str) -> Path:
    """Файл комплекта: сначала bundle/, иначе корень (обратная совместимость)."""
    p = BUNDLE_DIR / name
    if p.is_file():
        return p
    return ROOT / name


# Имена как в employees_io — копируются рядом с exe, если есть в корне проекта.
_EMPLOYEES_XLSX = "Data_base.xlsx"
_PROGRAMS_XLSX = "Programs_base.xlsx"
NEXT = ROOT / "ProtocolOHT_next"
DEFAULT_OUT_DIR = ROOT / "ProtocolOHT_onefile"
EXE_NAME = "ProtocolOOT"
WORK = ROOT / "_pyinstaller_build_onefile"

# main.py: сначала ProtocolOHT_next в sys.path, затем корень. Для PyInstaller тот же порядок --paths.
# В корне проекта не дублируйте модули из ProtocolOHT_next (protocol_docx, protocol_ui, protocol_paths,
# protocol_recovery) — PyInstaller иначе может упаковать устаревшую копию и получить ImportError в exe.
_PYI_HIDDEN = [
    "app_paths",
    "program_keys",
    "clipboard_ui",
    "commission_admin",
    "employees_io",
    "excel_data_cache",
    "docx_template_protection",
    "programs_v_prof",
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

# Копируются в data/ рядом с exe (.md в поставку не включаются — только docx/xlsx/txt).
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
# Образцы Excel — дополнительно в корень с exe (программа подхватывает их при первом запуске).
BUNDLE_EXE_ROOT_XLSX = (_EMPLOYEES_XLSX, _PROGRAMS_XLSX)


def _pick_output_dir_interactive() -> Path | None:
    """Диалог выбора папки; None — пользователь отменил или tkinter недоступен."""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        print(
            "tkinter недоступен — используется папка по умолчанию:",
            DEFAULT_OUT_DIR,
            file=sys.stderr,
        )
        return DEFAULT_OUT_DIR
    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
    except tk.TclError:
        pass
    path = filedialog.askdirectory(
        title="Папка для сборки: сюда попадут ProtocolOOT.exe и шаблоны",
        initialdir=str(DEFAULT_OUT_DIR.parent),
    )
    root.destroy()
    if not path:
        return None
    return Path(path).expanduser().resolve()


DIST_README = """Папка готовой сборки (onefile + комплект в data/)

• ProtocolOOT.exe — программа (один файл).
• Подпапка data/ — шаблоны Word, образцы Excel, XSD Минтруда, справка FAQ.txt, инструкции (.docx).
  В корне с exe эти файлы не лежат — так проще не путать их с рабочими базами.

При сборке Data_base.xlsx и Programs_base.xlsx (если есть в исходниках) кладутся в data/ и
дублируются в корень рядом с exe. Без них при первом запуске в корне создаётся пустой шаблон
сотрудников — для протокола Word нужны заполненные Excel или свои файлы в настройках.

При работе рядом с exe появятся protocols.db, last_protocol_no.json, журнал ошибок,
папки Protokol и Mintrud (рабочие данные — в корне, не в data/).

Переносите на другой ПК всю папку: exe + data/ целиком.
В корне рядом с exe лежат копии Data_base.xlsx и Programs_base.xlsx (если были в исходниках при сборке).

Для PDF с оформлением Word на целевом ПК нужны Microsoft Word и регистрация COM (pywin32 входит в сборку exe).
"""


def _run_verify_project() -> int:
    """Импорты и ruff перед PyInstaller (tools/verify_project.py --no-launch)."""
    script = ROOT / "tools" / "verify_project.py"
    if not script.is_file():
        return 0
    print("Проверка проекта (verify_project.py --no-launch)...")
    return subprocess.run(
        [sys.executable, str(script), "--no-launch"],
        cwd=str(ROOT),
    ).returncode


def _run_ruff_check() -> int:
    """Перед сборкой: ruff по ruff.toml (F821 и др.) — ловит пропущенные импорты и синтаксис."""
    rc = subprocess.run(
        [sys.executable, "-m", "ruff", "check", "."],
        cwd=str(ROOT),
    ).returncode
    if rc != 0:
        print(
            "Сборка остановлена: исправьте замечания ruff. Установка: py -3 -m pip install ruff",
            file=sys.stderr,
        )
    return rc


def main() -> int:
    os.chdir(ROOT)
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("Установите PyInstaller: py -3 -m pip install -r requirements-build.txt", file=sys.stderr)
        return 1

    if _run_verify_project() != 0:
        return 1
    print("Ruff: проверка исходников (ruff.toml)...")
    if _run_ruff_check() != 0:
        return 1

    if len(sys.argv) > 1:
        out_arg = " ".join(sys.argv[1:]).strip().strip('"')
        OUT_DIR = Path(out_arg).expanduser().resolve()
    else:
        picked = _pick_output_dir_interactive()
        if picked is None:
            print("Папка не выбрана — сборка отменена.")
            return 0
        OUT_DIR = picked

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
        f"--specpath={ROOT}",
        f"--paths={NEXT}",
        f"--paths={ROOT}",
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

    args.append(str(ROOT / "main.py"))

    print("PyInstaller: --onefile, выход:", OUT_DIR)
    r = subprocess.run(args, cwd=str(ROOT))
    if r.returncode != 0:
        return r.returncode

    exe = OUT_DIR / f"{EXE_NAME}.exe"
    if not exe.is_file():
        print("Ошибка: не найден", exe, file=sys.stderr)
        return 1

    data_dir = OUT_DIR / DATA_SUBDIR
    data_dir.mkdir(parents=True, exist_ok=True)

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

    for label, xname in (
        ("база сотрудников", _EMPLOYEES_XLSX),
        ("справочник программ", _PROGRAMS_XLSX),
    ):
        if not _bundle_src(xname).is_file() and not (ROOT / xname).is_file():
            print(
                f"  Внимание: нет {xname} (корень или bundle/) — в {DATA_SUBDIR}/ {label} не скопирован.",
                file=sys.stderr,
            )

    (OUT_DIR / "ИНСТРУКЦИЯ_папки_сборки.txt").write_text(DIST_README, encoding="utf-8")

    print()
    print("Сборка завершена.")
    print(f"  {exe}")
    print(f"  Комплект в {data_dir.name}/: {copied} файл(ов)")
    print()
    print(f"Переносите на другие ПК всю папку:\n  {OUT_DIR}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
