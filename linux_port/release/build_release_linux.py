# -*- coding: utf-8 -*-
"""Сборка релиза ProtocolOOT для Linux (PyInstaller + комплект файлов).

По образцу grafik-pz/release/build_release_linux.py:
  python linux_port/release/build_release_linux.py --local
  python linux_port/release/build_release_linux.py --binary-only
  python linux_port/release/build_release_linux.py -o /path/to/out
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

LINUX_PORT = Path(__file__).resolve().parents[1]
APP = LINUX_PORT / "app"
BUNDLE_DIR = APP / "bundle"
SPEC_PATH = LINUX_PORT / "release" / "protocol_oot_linux.spec"
BUILD_WORK_DIR = LINUX_PORT / "release" / "_build"
PYINSTALLER_DIST = LINUX_PORT / "dist"
DEFAULT_LOCAL_OUTPUT = LINUX_PORT / "release" / "out_linux"
BINARY_NAME = "ProtocolOOT"
INSTRUCTION_NAME = "ИНСТРУКЦИЯ_папки_сборки.txt"

_EMPLOYEES_XLSX = "Data_base.xlsx"
_PROGRAMS_XLSX = "Programs_base.xlsx"
BUNDLE_FILES = [
    "default_protocol.docx",
    "default_protocol_tehnicheskiy.docx",
    "ПОДРОБНАЯ_ИНСТРУКЦИЯ_для_пользователя.docx",
    "ИНСТРУКЦИЯ_оформление_протоколов_Минтруд.docx",
    "ЖУРНАЛ_ДОРАБОТОК.md",
    _EMPLOYEES_XLSX,
    _PROGRAMS_XLSX,
    "icon.ico",
    "Шаблон_Минтруд_XSD_УМН.xlsx",
    "!! Шаблон_Минтруд_XSD_УМН _ общ+.xlsx",
    "Шаблон_Минтруд_XSD_УМН _ общ+.xlsx",
]
BUNDLE_EXE_ROOT_XLSX = (_EMPLOYEES_XLSX, _PROGRAMS_XLSX)


def _bundle_src(name: str) -> Path:
    p = BUNDLE_DIR / name
    if p.is_file():
        return p
    return APP / name


def _data_subdir_name() -> str:
    sys.path.insert(0, str(APP / "ProtocolOHT_next"))
    sys.path.insert(1, str(APP))
    from app_paths import application_resource_data_subdir_name  # noqa: E402

    return application_resource_data_subdir_name()


def ensure_pyinstaller() -> None:
    req = LINUX_PORT / "requirements-build.txt"
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", str(req)])
        return
    # verify/PyInstaller нужны и pip-зависимости приложения (openpyxl, docx, …)
    try:
        import openpyxl  # noqa: F401
        import docx  # noqa: F401
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", str(req)])


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
    sys.path.insert(0, str(LINUX_PORT))
    from ruff_linux import run_ruff_on_app

    rc = run_ruff_on_app(linux_port=LINUX_PORT, app_dir=APP, cwd=LINUX_PORT)
    if rc != 0:
        print("Сборка остановлена: исправьте замечания ruff.", file=sys.stderr)
    return rc


def run_pyinstaller() -> Path:
    if not (APP / "main.py").is_file():
        msg = f"Нет {APP}/main.py — выполните: python linux_port/prepare.py"
        raise FileNotFoundError(msg)

    BUILD_WORK_DIR.mkdir(parents=True, exist_ok=True)
    PYINSTALLER_DIST.mkdir(parents=True, exist_ok=True)
    subprocess.check_call(
        [
            sys.executable,
            "-m",
            "PyInstaller",
            "--noconfirm",
            "--clean",
            "--distpath",
            str(PYINSTALLER_DIST),
            "--workpath",
            str(BUILD_WORK_DIR / "work_linux"),
            str(SPEC_PATH),
        ],
        cwd=str(LINUX_PORT),
    )
    binary = PYINSTALLER_DIST / BINARY_NAME
    if binary.is_file():
        return binary
    candidates = [p for p in PYINSTALLER_DIST.iterdir() if p.is_file()]
    if len(candidates) == 1:
        return candidates[0]
    msg = f"Бинарник {binary} не найден после сборки. Проверьте {PYINSTALLER_DIST}."
    raise FileNotFoundError(msg)


def _format_size(path: Path) -> str:
    size_mb = path.stat().st_size / 1024 / 1024
    return f"{size_mb:.1f} МБ"


def assemble_release(output_dir: Path, *, binary_source: Path) -> dict[str, Path]:
    """Бинарник, инструкция и каталог data/ в выбранном каталоге."""
    output_dir = output_dir.resolve()
    data_subdir = _data_subdir_name()

    if output_dir.exists():
        for item in output_dir.iterdir():
            if item.is_file() and item.name in {BINARY_NAME, INSTRUCTION_NAME, *BUNDLE_EXE_ROOT_XLSX}:
                item.unlink()
            elif item.is_dir() and item.name == data_subdir:
                shutil.rmtree(item)
    output_dir.mkdir(parents=True, exist_ok=True)

    binary_dest = output_dir / BINARY_NAME
    shutil.copy2(binary_source, binary_dest)
    binary_dest.chmod(binary_dest.stat().st_mode | 0o111)

    instruction_src = SPEC_PATH.parent / INSTRUCTION_NAME
    instruction_dest = output_dir / INSTRUCTION_NAME
    shutil.copy2(instruction_src, instruction_dest)

    data_dir = output_dir / data_subdir
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
                shutil.copy2(src, output_dir / name)

    faq_src = _bundle_src("FAQ.md")
    if faq_src.is_file():
        shutil.copy2(faq_src, data_dir / "FAQ.txt")
        copied += 1

    (data_dir / "README.txt").write_text(
        "Каталог комплекта программы: шаблоны, справка, образцы Excel.\n"
        "protocols.db и рабочие папки создаются рядом с ProtocolOOT при работе.\n",
        encoding="utf-8",
    )

    return {
        "binary": binary_dest,
        "instruction": instruction_dest,
        "data": data_dir,
    }


def build_release(output_dir: Path, *, skip_verify: bool = False) -> dict[str, Path]:
    ensure_pyinstaller()
    if not skip_verify:
        if _run_verify() != 0:
            raise SystemExit(1)
        if _run_ruff() != 0:
            raise SystemExit(1)
    binary = run_pyinstaller()
    return assemble_release(output_dir, binary_source=binary)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Собрать ProtocolOOT для Linux (бинарник + data/)",
    )
    parser.add_argument("-o", "--output", type=Path, help="Папка для комплекта")
    parser.add_argument(
        "--local",
        action="store_true",
        help=f"Собрать комплект в {DEFAULT_LOCAL_OUTPUT.relative_to(LINUX_PORT.parent)}",
    )
    parser.add_argument(
        "--binary-only",
        action="store_true",
        help="Только PyInstaller → linux_port/dist/ (без комплекта)",
    )
    parser.add_argument(
        "--no-verify",
        action="store_true",
        help="Без verify_linux и ruff (CI)",
    )
    args = parser.parse_args()

    if args.binary_only:
        ensure_pyinstaller()
        if not args.no_verify:
            if _run_verify() != 0:
                return 1
            if _run_ruff() != 0:
                return 1
        binary = run_pyinstaller()
        print(f"Бинарник: {binary} ({_format_size(binary)})")
        return 0

    output = args.output or (DEFAULT_LOCAL_OUTPUT if args.local else None)
    if output is None:
        print("Укажите --output или --local.", file=sys.stderr)
        return 1

    result = build_release(output, skip_verify=args.no_verify)
    print(f"Готово: {result['binary']} ({_format_size(result['binary'])})")
    print(f"Инструкция: {result['instruction']}")
    print(f"Data: {result['data']}")
    print(f"\nПереносите на другие ПК всю папку:\n  {output.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
