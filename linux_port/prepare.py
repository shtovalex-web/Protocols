# -*- coding: utf-8 -*-
"""
Подготовка Linux-копии приложения из основного проекта (без правок в корне репозитория).

Запуск из корня проекта или из linux_port/:
    python linux_port/prepare.py
    python prepare.py   # из каталога linux_port

Результат: linux_port/app/ — готовая к переносу на Linux копия с оверлеями.
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

LINUX_PORT = Path(__file__).resolve().parent
PROJECT_ROOT = LINUX_PORT.parent
APP_DIR = LINUX_PORT / "app"
OVERLAYS = LINUX_PORT / "overlays"

# Копируемые каталоги (относительно корня проекта).
COPY_DIRS = (
    "ProtocolOHT_next",
    "bundle",
)

# Копируемые файлы в корне app/.
COPY_ROOT_FILES = (
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
    "russian_genitive.py",
    "v_prof_combinations.py",
    "v_program_registry_match.py",
    "educated_person_import_v1.0.9.xsd",
)

# Текстовые замены в скопированных файлах (Linux-подсказки в UI).
TEXT_REPLACEMENTS: tuple[tuple[str, str, str], ...] = (
    (
        "ProtocolOHT_next/protocol_ui.py",
        (
            '"Не удалось сохранить PDF через Microsoft Word (нужен установленный Word и пакет pywin32, "\n'
            '                        "либо сбой при конвертации).\\n\\n"'
        ),
        (
            '"Не удалось сохранить PDF с оформлением DOCX (нужны LibreOffice "\n'
            '                        "или Microsoft Word + docx2pdf, либо сбой конвертации).\\n\\n"'
        ),
    ),
)

SKIP_DIR_NAMES = {
    "__pycache__",
    ".git",
    ".venv",
    "venv",
    "linux_port",
    "эталон_сборки",
    "ProtocolOHT_onefile",
    "ProtocolOHT_app",
    "_pyinstaller_build",
    "_pyinstaller_build_onefile",
    "Protokol",
    "Mintrud",
    "local",
    "ib_review",
}


def _should_skip(path: Path) -> bool:
    return any(part in SKIP_DIR_NAMES for part in path.parts)


def _copy_tree(src: Path, dst: Path) -> None:
    if not src.is_dir():
        return
    for item in src.rglob("*"):
        if _should_skip(item.relative_to(src)):
            continue
        rel = item.relative_to(src)
        target = dst / rel
        if item.is_dir():
            target.mkdir(parents=True, exist_ok=True)
        else:
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(item, target)


def _apply_overlays() -> int:
    if not OVERLAYS.is_dir():
        return 0
    count = 0
    for src in OVERLAYS.rglob("*"):
        if src.is_dir():
            continue
        rel = src.relative_to(OVERLAYS)
        dst = APP_DIR / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        count += 1
    return count


def _apply_text_replacements() -> int:
    changed = 0
    for rel, old, new in TEXT_REPLACEMENTS:
        path = APP_DIR / rel
        if not path.is_file():
            continue
        text = path.read_text(encoding="utf-8")
        if old not in text:
            continue
        path.write_text(text.replace(old, new, 1), encoding="utf-8")
        changed += 1
    return changed


def _copy_port_meta() -> None:
    for name in ("requirements.txt", "requirements-build.txt", "README_APP.md"):
        src = LINUX_PORT / name
        if src.is_file():
            shutil.copy2(src, APP_DIR / name)


def prepare() -> int:
    if not (PROJECT_ROOT / "main.py").is_file():
        print(f"Ошибка: не найден основной проект в {PROJECT_ROOT}", file=sys.stderr)
        return 1

    if APP_DIR.exists():
        shutil.rmtree(APP_DIR)
    APP_DIR.mkdir(parents=True)

    for dirname in COPY_DIRS:
        src = PROJECT_ROOT / dirname
        if src.is_dir():
            _copy_tree(src, APP_DIR / dirname)

    for name in COPY_ROOT_FILES:
        src = PROJECT_ROOT / name
        if src.is_file():
            shutil.copy2(src, APP_DIR / name)

    overlays = _apply_overlays()
    patches = _apply_text_replacements()
    _copy_port_meta()

    print(f"Linux-копия подготовлена: {APP_DIR}")
    print(f"  Оверлеев: {overlays}")
    print(f"  Текстовых патчей UI: {patches}")
    print()
    print("Дальше на Linux:")
    print("  cd linux_port && ./install_deps.sh && ./run.sh")
    print("  cd linux_port && ./build_linux.sh")
    return 0


if __name__ == "__main__":
    sys.exit(prepare())
