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
    py -3 build_windows_exe.py --print-version
    или двойной щелчок по build_windows_exe.bat (папка DEPLOY_ROOT\\<версия>\\, см. bat)

Без аргументов сначала открывается диалог выбора папки (tkinter); «Отмена» — выход без сборки.
Необязательный аргумент — путь к папке вывода (перетаскивание на .bat, создание при необходимости).

Если PyInstaller не ставится на очень новый Python (например 3.14), соберите на ПК с Python 3.11–3.12.
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from datetime import date
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
DEPLOY_UPDATE_SHARE_DIR = Path(r"D:\Обновление")
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
    "startup_update",
    "update_config",
    "update_manifest",
    "update_installer",
    "update_scan",
    "update_bundle_files",
    "update_data_installer",
    "update_info",
    "version_compare",
    "changelog_dialog",
    "pending_changelog",
    "windows_app_bundle",
    "update_success",
    "protocol_updater_config",
    "desktop_updater",
    "desktop_updater.config",
    "desktop_updater.registry",
    "desktop_updater.client_config",
    "desktop_updater.manifest",
    "desktop_updater.scan",
    "desktop_updater.installer",
    "desktop_updater.data_installer",
    "desktop_updater.info",
    "desktop_updater.bundle_files",
    "desktop_updater.version_compare",
    "desktop_updater.pending_changelog",
    "desktop_updater.success",
    "desktop_updater.startup",
    "desktop_updater.windows_app_bundle",
]

# fpdf2 тянет fontTools; на Python 3.14 iup — бинарный .pyd, без collect/hidden-import exe падает при старте.
_PYI_COLLECT_SUBMODULES = ("openpyxl", "pymorphy3", "pymorphy2", "fontTools")
_PYI_EXTRA_HIDDEN = (
    "docx",
    "docx.oxml",
    "fpdf.fonts",
    "fpdf.enums",
    "fontTools",
    "fontTools.varLib",
    "fontTools.varLib.iup",
    "_overlapped",
    "asyncio.windows_events",
)

# Копируются в data/ рядом с exe (.md в поставку не включаются — только docx/xlsx/txt).
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
# Образцы Excel — дополнительно в корень с exe (программа подхватывает их при первом запуске).
BUNDLE_EXE_ROOT_XLSX = (_EMPLOYEES_XLSX, _PROGRAMS_XLSX)


def _copy_bundle_asset(src: Path, dst: Path) -> bool:
    """Копирует файл комплекта; False — файл занят или нет прав (сборка exe уже готова)."""
    try:
        shutil.copy2(src, dst)
        return True
    except OSError as error:
        print(f"  Внимание: не удалось скопировать {dst.name}: {error}", file=sys.stderr)
        return False


def _ensure_update_config(out_dir: Path, *, share_root: Path | None = None) -> bool:
    """Создаёт update_config.json с путём к шаре; существующий файл не трогает."""
    path = out_dir / "update_config.json"
    if path.is_file():
        return False
    if str(NEXT) not in sys.path:
        sys.path.insert(0, str(NEXT))
    from update_config import DEFAULT_UPDATE_SHARE_ROOT, format_manifest_path_for_json

    target_share = share_root
    if target_share is None:
        target_share = (
            DEPLOY_UPDATE_SHARE_DIR
            if out_dir.resolve() == DEFAULT_OUT_DIR.resolve()
            else DEFAULT_UPDATE_SHARE_ROOT
        )
    payload = {
        "manifest_path": format_manifest_path_for_json(target_share),
        "enabled": True,
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"  Создан {path.name} -> {target_share}")
    return True


def _load_publish_module():
    import importlib.util

    publish_script = ROOT / "tools" / "publish_update_manifest.py"
    spec = importlib.util.spec_from_file_location("publish_update_manifest", publish_script)
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(mod)
    return mod


def _write_update_info(data_dir: Path) -> Path:
    """Записывает data/update_info.json с версией текущей сборки."""
    if str(NEXT) not in sys.path:
        sys.path.insert(0, str(NEXT))
    from protocol_app_info import APP_VERSION
    from update_info import write_update_info

    version = (APP_VERSION or "").strip()
    if not version:
        msg = "APP_VERSION пуст — не удалось записать update_info.json"
        raise ValueError(msg)
    return write_update_info(data_dir, version=version, released=date.today().isoformat())


def _publish_update_share(
    *,
    exe: Path,
    data_dir: Path,
    share_root: Path,
    app_zip: Path | None = None,
) -> Path:
    """Публикует exe + data/ на шару обновлений (windows/<версия>/…)."""
    if str(NEXT) not in sys.path:
        sys.path.insert(0, str(NEXT))
    from protocol_app_info import APP_VERSION

    version = (APP_VERSION or "").strip()
    if not version:
        msg = f"APP_VERSION пуст — не удалось опубликовать в {share_root}"
        raise ValueError(msg)
    share_root.mkdir(parents=True, exist_ok=True)
    publish = _load_publish_module().publish
    return publish(
        exe_path=exe,
        version=version,
        share_root=share_root,
        changes=[f"Сборка {version}"],
        mandatory=False,
        released=date.today().isoformat(),
        data_src_dir=data_dir,
        app_zip=app_zip,
    )


def _write_release_manifest(
    *,
    exe: Path,
    data_dir: Path,
    release_dir: Path,
) -> Path:
    """manifest.json в папке релиза (рядом с exe и data/)."""
    if str(NEXT) not in sys.path:
        sys.path.insert(0, str(NEXT))
    from protocol_app_info import APP_VERSION

    version = (APP_VERSION or "").strip()
    if not version:
        msg = "APP_VERSION пуст — не удалось записать manifest.json в папку релиза"
        raise ValueError(msg)
    publish = _load_publish_module()
    return publish._write_manifest(
        target_dir=release_dir,
        target_payload=exe,
        target_data=data_dir,
        version=version,
        changes=[f"Сборка {version}"],
        mandatory=False,
        released=date.today().isoformat(),
    )


def _try_publish_deploy_update_share(
    *,
    exe: Path,
    data_dir: Path,
    app_zip: Path | None = None,
) -> Path | None:
    """Публикует на D:\\Обновление\\windows\\<версия>/; при ошибке — предупреждение."""
    try:
        return _publish_update_share(
            exe=exe,
            data_dir=data_dir,
            share_root=DEPLOY_UPDATE_SHARE_DIR,
            app_zip=app_zip,
        )
    except (OSError, ValueError, SystemExit) as error:
        print(
            f"\nВнимание: не удалось опубликовать в {DEPLOY_UPDATE_SHARE_DIR}: {error}",
            file=sys.stderr,
        )
        return None


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
• Подпапка data/ — шаблоны Word, образцы Excel, XSD Минтруда, справка FAQ.txt,
  инструкции (.docx), журнал доработок (ЖУРНАЛ_ДОРАБОТОК.md).
  В корне с exe эти файлы не лежат — так проще не путать их с рабочими базами.

При сборке Data_base.xlsx и Programs_base.xlsx (если есть в исходниках) кладутся в data/ и
дублируются в корень рядом с exe. Без них при первом запуске в корне создаётся пустой шаблон
сотрудников — для протокола Word нужны заполненные Excel или свои файлы в настройках.

При работе рядом с exe появятся protocols.db, last_protocol_no.json, журнал ошибок,
папки Protokol и Mintrud (рабочие данные — в корне, не в data/).

Переносите на другой ПК всю папку: exe + data/ целиком.
Избегайте символа «!» в пути к папке — он мешает перезапуску после обновления в Windows.
В корне рядом с exe лежат копии Data_base.xlsx и Programs_base.xlsx (если были в исходниках при сборке).

update_config.json — каталог шары обновлений (создаётся при первой сборке, если файла ещё нет;
существующий не перезаписывается). Без файла — \\\\tn.tngrp.ru\\df\\AK\\SHR\\Distr_О_О\\Шитов Алексей Александрович из кода программы.

manifest.json создаётся в этой папке при сборке (рядом с exe) и дублируется на шаре в windows/<версия>/manifest.json.
После сборки комплект для обновления публикуется на шару (D:\\Обновление при сборке в ProtocolOHT_onefile или UNC выше).
В data/ записывается update_info.json — маркер версии комплекта шаблонов и справки.
Для ProtocolOHT_onefile update_config.json по умолчанию указывает на D:/Обновление (если файла ещё нет).

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


def _app_version() -> str:
    if str(NEXT) not in sys.path:
        sys.path.insert(0, str(NEXT))
    from protocol_app_info import APP_VERSION

    return (APP_VERSION or "").strip()


def main() -> int:
    os.chdir(ROOT)
    if len(sys.argv) == 2 and sys.argv[1] == "--print-version":
        version = _app_version()
        if not version:
            print("?", file=sys.stderr)
            return 1
        print(version)
        return 0

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

    if sys.version_info >= (3, 14):
        print(
            "\nВНИМАНИЕ: сборка на Python 3.14+ — PyInstaller onefile может падать "
            'с «Failed to load Python DLL» при перезапуске.\n'
            "Рекомендуется: py -3.12 build_windows_exe.py …\n",
            file=sys.stderr,
        )

    try:
        import pymorphy3  # noqa: F401
        import pymorphy3_dicts_ru  # noqa: F401
    except ImportError as e:
        print(
            "Ошибка: для родительного падежа комиссии в .exe нужны pymorphy3 и словари.\n"
            f"  {e}\n"
            "Установите в тот же Python, которым собираете (обычно 3.12):\n"
            "  py -3.12 -m pip install pymorphy3 pymorphy3-dicts-ru\n"
            "  или: py -3.12 -m pip install -r requirements.txt",
            file=sys.stderr,
        )
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
    for mod in _PYI_EXTRA_HIDDEN:
        args.append(f"--hidden-import={mod}")
    for pkg in _PYI_COLLECT_SUBMODULES:
        args.append(f"--collect-submodules={pkg}")
    args.append("--collect-all=tkinter")
    args.append(f"--runtime-hook={ROOT / 'tools' / 'pyi_rth_tkinter.py'}")
    # Проверено выше: pymorphy3 + словари обязательны для сборки.
    args.append("--collect-data=pymorphy3_dicts_ru")
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
    copy_failures: list[str] = []
    for name in BUNDLE_FILES:
        src = _bundle_src(name)
        if src.is_file():
            if _copy_bundle_asset(src, data_dir / name):
                copied += 1
                if name in BUNDLE_EXE_ROOT_XLSX:
                    _copy_bundle_asset(src, OUT_DIR / name)
            else:
                copy_failures.append(name)

    faq_src = _bundle_src("FAQ.md")
    if faq_src.is_file():
        if _copy_bundle_asset(faq_src, data_dir / "FAQ.txt"):
            copied += 1
        else:
            copy_failures.append("FAQ.txt")

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
    dev_build = OUT_DIR.resolve() == DEFAULT_OUT_DIR.resolve()
    try:
        update_info_path = _write_update_info(data_dir)
    except (OSError, ValueError) as error:
        print(f"\nВнимание: не удалось записать update_info.json: {error}", file=sys.stderr)
        return 1

    _ensure_update_config(
        OUT_DIR,
        share_root=DEPLOY_UPDATE_SHARE_DIR if dev_build else None,
    )

    release_manifest: Path | None = None
    if not copy_failures:
        try:
            release_manifest = _write_release_manifest(
                exe=exe,
                data_dir=data_dir,
                release_dir=OUT_DIR,
            )
        except (OSError, ValueError, SystemExit) as error:
            print(
                f"\nВнимание: не удалось записать manifest.json в папку релиза: {error}",
                file=sys.stderr,
            )

    deploy_manifest = _try_publish_deploy_update_share(exe=exe, data_dir=data_dir)

    print()
    print("Сборка завершена.")
    print(f"  {exe}")
    print(f"  Комплект в {data_dir.name}/: {copied} файл(ов)")
    print(f"  Маркер версии: {update_info_path}")
    if release_manifest is not None:
        print(f"  manifest релиза: {release_manifest}")
    if deploy_manifest is not None:
        print(f"  Шара обновлений: {DEPLOY_UPDATE_SHARE_DIR}")
        print(f"  manifest: {deploy_manifest}")
    if copy_failures:
        print(
            f"\nВнимание: не скопированы {len(copy_failures)} файл(ов): "
            + ", ".join(copy_failures),
            file=sys.stderr,
        )
        print(
            "Закройте ProtocolOOT.exe, Word и повторите только копирование "
            f"(или пересоберите). Папка: {data_dir}",
            file=sys.stderr,
        )
        return 1
    print()
    print(f"Переносите на другие ПК всю папку:\n  {OUT_DIR}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
