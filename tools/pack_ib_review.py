# -*- coding: utf-8 -*-
"""
Архив исходников для проверки ИБ: код, bundle, тесты, Linux-порт, без пользовательских данных.

    py -3 tools/pack_ib_review.py
    py -3 tools/pack_ib_review.py -o D:/temp/ib_review.zip

Создаёт ib_review_YYYYMMDD_HHMMSS.zip в корне проекта (или путь из -o).
"""

from __future__ import annotations

import argparse
import os
import shutil
import stat
import subprocess
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable

ROOT = Path(__file__).resolve().parent.parent
BUNDLE = ROOT / "bundle"
NEXT = ROOT / "ProtocolOHT_next"
TOOLS = ROOT / "tools"

ROOT_PY = (
    "main.py",
    "app_paths.py",
    "bundle_integration.py",
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
    ".gitattributes",
    "build_windows_exe.py",
    "build_windows_exe.bat",
    "verify.bat",
    "update_etalon.bat",
    "sync_github.bat",
    "sync_linux_branch.bat",
    "generate_oformlenie_instruction_docx.bat",
    "generate_podrobnaya_instruction_docx.bat",
)

TOOLS_PY = (
    "verify_project.py",
    "update_etalon.py",
    "pack_ib_review.py",
    "pack_linux_build.py",
    "sync_github.py",
    "sync_linux_branch.py",
    "sync_linux_local.py",
    "tidy_workspace.py",
    "instruction_md_to_docx.py",
    "generate_oformlenie_instruction_docx.py",
    "generate_podrobnaya_instruction_docx.py",
    "_gen_embedded_png.py",
    "repair_technical_template_docx.py",
    "repair_protocol_templates_docx.py",
    "protect_protocol_templates.py",
    "patch_protocol_template_markers.py",
    "remove_logo_background.py",
    "capture_manual_screenshots.py",
)

LINUX_PORT_SKIP_PARTS = frozenset(
    {
        "app",
        ".venv",
        ".venv-linux",
        "__pycache__",
        "_build",
        "out_linux",
        "out_linux.zip",
        "ProtocolOHT_linux_dist",
        "_pyinstaller_build_linux",
        "ProtocolOOT.spec",
    }
)

IB_README = """Пакет исходников ProtocolOOT для проверки информационной безопасности
================================================================================

Состав: исходный код, шаблоны, тесты, скрипты сборки Windows/Linux.

НЕ включено (персональные / рабочие данные и артефакты):
  • protocols.db, last_protocol_no.json, protocol_errors_journal.txt
  • пользовательские Excel и протоколы (Protokol/, Mintrud/, local/)
  • эталон_сборки/, собранные exe и zip-сборки
  • linux_port/app/ (генерируется prepare.py), venv, кэши, .git, .cursor

Включено:
  • Python: корень + ProtocolOHT_next/
  • bundle/ — шаблоны, FAQ, инструкции, XSD Минтруда
  • tests/ — unit-тесты
  • linux_port/ — сборка под Linux (без сгенерированной app/)
  • docs/ — структура проекта и сборка
  • tools/ — проверка, упаковка, генерация docx
  • Data_base.xlsx / Programs_base.xlsx — пустые шаблоны (без сотрудников)

Запуск из исходников:
  py -3 -m pip install -r requirements.txt
  py -3 main.py

Проверка:
  py -3 tools/verify_project.py --no-launch

Сборка exe (Windows):
  py -3 -m pip install -r requirements-build.txt
  py -3 build_windows_exe.py
"""


def _rmtree_resilient(path: Path) -> None:
    """Удалить каталог; на Windows снимает read-only у защищённых docx."""

    def onexc(func, p, exc_info):
        exc = exc_info[1] if isinstance(exc_info, tuple) else exc_info
        if isinstance(exc, PermissionError):
            os.chmod(p, stat.S_IWRITE)
            func(p)
            return
        raise exc

    shutil.rmtree(path, onexc=onexc)


def _copy_file(src: Path, dest: Path) -> bool:
    if not src.is_file():
        return False
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dest)
    return True


def _linux_port_skip(rel: Path) -> bool:
    parts = rel.parts
    if not parts:
        return False
    if parts[0] in LINUX_PORT_SKIP_PARTS:
        return True
    if "release" in parts:
        idx = parts.index("release")
        if idx + 1 < len(parts) and parts[idx + 1] in LINUX_PORT_SKIP_PARTS:
            return True
    return False


def _copy_tree(
    src_root: Path,
    dest_root: Path,
    *,
    skip: Callable[[Path], bool] | None = None,
) -> list[str]:
    copied: list[str] = []
    if not src_root.is_dir():
        return copied
    for src in src_root.rglob("*"):
        if not src.is_file():
            continue
        rel = src.relative_to(src_root)
        if skip and skip(rel):
            continue
        if rel.name.endswith((".pyc", ".pyo")) or "__pycache__" in rel.parts:
            continue
        dest = dest_root / rel
        _copy_file(src, dest)
        copied.append(f"{dest_root.name}/{rel.as_posix()}")
    return copied


def _zip_dir(folder: Path, zip_path: Path) -> None:
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _dirs, files in os.walk(folder):
            for name in files:
                full = Path(root) / name
                arc = full.relative_to(folder).as_posix()
                zf.write(full, arc)


def _git_head() -> str:
    try:
        r = subprocess.run(
            ["git", "rev-parse", "--short", "HEAD"],
            cwd=str(ROOT),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if r.returncode == 0:
            return (r.stdout or "").strip()
    except OSError:
        pass
    return ""


def _app_version() -> str:
    try:
        info = ROOT / "ProtocolOHT_next" / "protocol_app_info.py"
        text = info.read_text(encoding="utf-8")
        m = __import__("re").search(r'APP_VERSION\s*=\s*"([^"]+)"', text)
        if m:
            return m.group(1)
    except OSError:
        pass
    return "?"


def pack(*, out_zip: Path | None = None) -> Path:
    os.chdir(ROOT)
    sys.path.insert(0, str(ROOT))

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    staging = ROOT / f"_ib_staging_{stamp}"
    zip_path = out_zip or (ROOT / f"ib_review_{stamp}.zip")
    zip_path = zip_path.expanduser().resolve()

    if staging.exists():
        _rmtree_resilient(staging)
    staging.mkdir(parents=True)

    copied: list[str] = []
    missing: list[str] = []

    for name in ROOT_PY + ROOT_MISC:
        if _copy_file(ROOT / name, staging / name):
            copied.append(name)
        elif name in ROOT_PY or name in ("README.md", "requirements.txt", "build_windows_exe.py"):
            missing.append(name)

    copied.extend(_copy_tree(BUNDLE, staging / "bundle"))
    copied.extend(_copy_tree(NEXT, staging / "ProtocolOHT_next"))

    for name in TOOLS_PY:
        if _copy_file(TOOLS / name, staging / "tools" / name):
            copied.append(f"tools/{name}")

    copied.extend(_copy_tree(ROOT / "docs", staging / "docs"))
    copied.extend(_copy_tree(ROOT / "tests", staging / "tests"))
    copied.extend(
        _copy_tree(ROOT / "linux_port", staging / "linux_port", skip=_linux_port_skip)
    )
    copied.extend(_copy_tree(ROOT / ".github", staging / ".github"))

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

    manifest = (
        f"ProtocolOOT — пакет для проверки ИБ\n"
        f"version={_app_version()}\n"
        f"packed_utc={datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')}\n"
        f"git_commit={_git_head() or 'n/a'}\n"
        f"files={len(copied)}\n"
    )
    (staging / "ИБ_ПАМЯТКА.txt").write_text(IB_README, encoding="utf-8")
    (staging / "ИБ_MANIFEST.txt").write_text(manifest, encoding="utf-8")
    copied.append("ИБ_ПАМЯТКА.txt")
    copied.append("ИБ_MANIFEST.txt")

    if missing:
        print("Не найдено (обязательное):", ", ".join(missing), file=sys.stderr)

    if zip_path.is_file():
        zip_path.unlink()
    _zip_dir(staging, zip_path)
    _rmtree_resilient(staging)
    return zip_path


def main() -> int:
    parser = argparse.ArgumentParser(description="Архив исходников для проверки ИБ")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Путь к zip (по умолчанию ib_review_YYYYMMDD_HHMMSS.zip в корне)",
    )
    args = parser.parse_args()
    zip_path = pack(out_zip=args.output)
    size_mb = zip_path.stat().st_size / (1024 * 1024)
    print(f"Готово: {zip_path}")
    print(f"Размер: {size_mb:.2f} МБ")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
