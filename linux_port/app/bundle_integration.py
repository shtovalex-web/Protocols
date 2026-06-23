# -*- coding: utf-8 -*-
"""Файлы комплекта bundle/: поиск .docx/.odt, .xlsx/.ods и конвертация для openpyxl/python-docx."""

from __future__ import annotations

import logging
import shutil
import subprocess
import sys
from pathlib import Path

from app_paths import application_bundle_dir, application_exe_dir, application_user_dir

_logger = logging.getLogger(__name__)

OPENPYXL_EXTENSIONS = frozenset({".xlsx", ".xlsm"})
WORD_TEMPLATE_EXTENSIONS = frozenset({".docx", ".docm"})
LIBREOFFICE_SPREADSHEET_EXTENSIONS = frozenset({".ods", ".xls"})
LIBREOFFICE_WORD_EXTENSIONS = frozenset({".odt"})


class BundleOfficeConvertError(Exception):
    """Не удалось прочитать ODS/ODT — нет LibreOffice или ошибка конвертации."""


def office_cache_dir() -> Path:
    p = application_user_dir() / ".office_cache"
    p.mkdir(parents=True, exist_ok=True)
    return p


def _libreoffice_soffice() -> str | None:
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    if sys.platform == "win32":
        for base in (
            Path(r"C:\Program Files\LibreOffice"),
            Path(r"C:\Program Files (x86)\LibreOffice"),
        ):
            exe = base / "program" / "soffice.exe"
            if exe.is_file():
                return str(exe)
    return None


def find_bundle_asset(stem: str, *extensions: str) -> Path | None:
    """Ищет файл в bundle/ и в каталоге программы (stem + расширение)."""
    ext_list = extensions or (".xlsx", ".docx")
    dirs: list[Path] = []
    for d in (application_bundle_dir(), application_exe_dir()):
        if d not in dirs:
            dirs.append(d)
    for folder in dirs:
        for ext in ext_list:
            if not ext.startswith("."):
                ext = f".{ext}"
            p = folder / f"{stem}{ext}"
            if p.is_file():
                return p.resolve()
    return None


def _cache_target(source: Path, target_ext: str) -> Path:
    ext = target_ext if target_ext.startswith(".") else f".{target_ext}"
    stamp = int(source.stat().st_mtime)
    safe = source.name.replace(" ", "_")
    return office_cache_dir() / f"{safe}.{stamp}{ext}"


def _convert_via_libreoffice(source: Path, target_ext: str) -> Path:
    src = source.expanduser().resolve()
    if not src.is_file():
        raise BundleOfficeConvertError(f"Файл не найден:\n{src}")
    ext = target_ext.lstrip(".").lower()
    cached = _cache_target(src, ext)
    if cached.is_file() and cached.stat().st_mtime >= src.stat().st_mtime:
        return cached
    lo = _libreoffice_soffice()
    if lo is None:
        raise BundleOfficeConvertError(
            f"Для файла «{src.name}» нужен LibreOffice (конвертация в .{ext}).\n"
            "Установите LibreOffice или положите рядом с программой версию "
            f"в формате .{ext} (например из bundle/ скопируйте .xlsx / .docx)."
        )
    out_dir = office_cache_dir()
    if cached.is_file():
        try:
            cached.unlink()
        except OSError:
            pass
    cmd = [
        lo,
        "--headless",
        "--norestore",
        "--convert-to",
        ext,
        "--outdir",
        str(out_dir),
        str(src),
    ]
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except OSError as e:
        raise BundleOfficeConvertError(f"Не удалось запустить LibreOffice:\n{e}") from e
    except subprocess.TimeoutExpired as e:
        raise BundleOfficeConvertError(f"Конвертация «{src.name}» заняла слишком много времени.") from e
    produced = out_dir / f"{src.stem}.{ext}"
    if proc.returncode != 0 or not produced.is_file():
        err = (proc.stderr or proc.stdout or "").strip()
        raise BundleOfficeConvertError(
            f"LibreOffice не сконвертировал «{src.name}» в .{ext}."
            + (f"\n{err}" if err else "")
        )
    if produced.resolve() != cached.resolve():
        if cached.is_file():
            cached.unlink()
        produced.replace(cached)
    return cached


def resolve_openpyxl_workbook_path(path: Path) -> Path:
    """Путь к таблице, которую может открыть openpyxl (.xlsx/.xlsm или кэш из .ods/.xls)."""
    p = Path(path).expanduser().resolve()
    if not p.is_file():
        return p
    suf = p.suffix.lower()
    if suf in OPENPYXL_EXTENSIONS:
        return p
    if suf in LIBREOFFICE_SPREADSHEET_EXTENSIONS:
        return _convert_via_libreoffice(p, "xlsx")
    return p


def resolve_docx_template_path(path: Path) -> Path:
    """Путь к шаблону Word для python-docx (.docx/.docm или кэш из .odt)."""
    p = Path(path).expanduser().resolve()
    if not p.is_file():
        return p
    suf = p.suffix.lower()
    if suf in WORD_TEMPLATE_EXTENSIONS:
        return p
    if suf == ".txt":
        return p
    if suf in LIBREOFFICE_WORD_EXTENSIONS:
        return _convert_via_libreoffice(p, "docx")
    return p


def bundle_protocol_template_path(*, technical: bool = False) -> Path:
    stem = "default_protocol_tehnicheskiy" if technical else "default_protocol"
    found = find_bundle_asset(stem, *WORD_TEMPLATE_EXTENSIONS, *LIBREOFFICE_WORD_EXTENSIONS)
    if found is not None:
        return resolve_docx_template_path(found)
    name = (
        "default_protocol_tehnicheskiy.docx"
        if technical
        else "default_protocol.docx"
    )
    return application_bundle_dir() / name


def bundle_spreadsheet_path(stem: str) -> Path | None:
    found = find_bundle_asset(stem, *OPENPYXL_EXTENSIONS, *LIBREOFFICE_SPREADSHEET_EXTENSIONS)
    if found is None:
        return None
    return resolve_openpyxl_workbook_path(found)


def bundle_mintrud_template_path() -> Path | None:
    for stem in (
        "Шаблон_Минтруд_XSD_УМН",
        "!! Шаблон_Минтруд_XSD_УМН _ общ+",
        "Шаблон_Минтруд_XSD_УМН _ общ+",
    ):
        found = find_bundle_asset(stem, *OPENPYXL_EXTENSIONS, *LIBREOFFICE_SPREADSHEET_EXTENSIONS)
        if found is not None:
            try:
                return resolve_openpyxl_workbook_path(found)
            except BundleOfficeConvertError:
                _logger.warning("Не удалось подготовить шаблон Минтруда %s", found, exc_info=True)
    return None
