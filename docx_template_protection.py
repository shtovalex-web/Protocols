# -*- coding: utf-8 -*-
"""
Защита .docx-шаблонов протокола от правки в Microsoft Word (режим «только чтение»).

Ограничение действует при открытии файла в Word; программа по-прежнему читает шаблон
в память и не перезаписывает его при формировании протокола.

settings.xml меняется только вставкой тега (regex), без пересборки XML — иначе Word
может сообщать о «нечитаемом содержимом».
"""

from __future__ import annotations

import re
import stat
import sys
import zipfile
from collections.abc import Iterable
from pathlib import Path

_SETTINGS_PART = "word/settings.xml"
_DOC_PROT_OPEN = '<w:documentProtection w:edit="readOnly" w:enforcement="1"/>'
_DOC_PROT_RE = re.compile(
    r"<w:documentProtection\b[^>]*/>",
    re.IGNORECASE,
)
_DOC_PROT_BLOCK_RE = re.compile(
    r"<w:documentProtection\b[^>]*>.*?</w:documentProtection>",
    re.IGNORECASE | re.DOTALL,
)

_STANDARD_TEMPLATE_NAMES = (
    "default_protocol.docx",
    "default_protocol_tehnicheskiy.docx",
)


def standard_protocol_template_paths(base_dir: Path) -> list[Path]:
    """Стандартные шаблоны в каталоге (корень, bundle/, data/)."""
    base = Path(base_dir)
    out: list[Path] = []
    for name in _STANDARD_TEMPLATE_NAMES:
        for folder in (base, base / "bundle", base / "data"):
            p = folder / name
            if p.is_file():
                out.append(p.resolve())
    seen: set[str] = set()
    unique: list[Path] = []
    for p in out:
        key = str(p).lower()
        if key not in seen:
            seen.add(key)
            unique.append(p)
    return unique


def discover_standard_protocol_template_paths(
    extra_paths: Iterable[Path | str] | None = None,
) -> list[Path]:
    """
    Все копии стандартных шаблонов: корень поставки, bundle/, каталог bundle приложения (.exe).
    """
    from app_paths import application_bundle_dir, application_exe_dir

    candidates: list[Path] = []
    for base in (application_exe_dir(), application_bundle_dir()):
        candidates.extend(standard_protocol_template_paths(base))
    if extra_paths:
        for raw in extra_paths:
            p = Path(raw).expanduser()
            if p.is_file():
                candidates.append(p.resolve())
    seen: set[str] = set()
    unique: list[Path] = []
    for p in candidates:
        key = str(p).lower()
        if key not in seen:
            seen.add(key)
            unique.append(p)
    return unique


def _decode_settings(data: bytes) -> str:
    for enc in ("utf-8", "utf-8-sig", "cp1251"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def _encode_settings(text: str, original: bytes) -> bytes:
    if original.startswith(b"\xef\xbb\xbf"):
        return b"\xef\xbb\xbf" + text.encode("utf-8")
    if b'\r\n<w:settings' in original or original.startswith(b"<?xml"):
        return text.encode("utf-8")
    return text.encode("utf-8")


def _patch_settings_xml(data: bytes, *, enable: bool) -> bytes:
    """Вставка/удаление documentProtection без ElementTree (сохраняет исходные xmlns)."""
    text = _decode_settings(data)
    text = _DOC_PROT_BLOCK_RE.sub("", text)
    text = _DOC_PROT_RE.sub("", text)
    if enable:
        if "</w:settings>" in text:
            text = text.replace("</w:settings>", _DOC_PROT_OPEN + "</w:settings>", 1)
        else:
            text = text.rstrip() + _DOC_PROT_OPEN
    return _encode_settings(text, data)


def is_docx_word_readonly_protected(path: Path) -> bool:
    path = Path(path)
    if not path.is_file():
        return False
    try:
        with zipfile.ZipFile(path, "r") as zf:
            data = zf.read(_SETTINGS_PART)
    except (KeyError, OSError, zipfile.BadZipFile):
        return False
    text = _decode_settings(data).lower()
    return "documentprotection" in text and (
        'w:edit="readonly"' in text or "enforcement" in text
    )


def _rewrite_docx_member(path: Path, member: str, new_data: bytes) -> None:
    path = Path(path)
    tmp = path.with_suffix(path.suffix + ".protect.tmp")
    if tmp.is_file():
        tmp.unlink()
    with zipfile.ZipFile(path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                payload = new_data if info.filename == member else zin.read(info.filename)
                zout.writestr(info, payload)
    tmp.replace(path)


def apply_docx_word_readonly_protection(
    path: Path,
    *,
    windows_readonly: bool = True,
) -> None:
    path = Path(path)
    # Сначала снимаем атрибут Windows, иначе tmp.replace() на Win32 падает с PermissionError.
    set_windows_readonly_flag(path, False)
    with zipfile.ZipFile(path, "r") as zf:
        settings = zf.read(_SETTINGS_PART)
    new_settings = _patch_settings_xml(settings, enable=True)
    _rewrite_docx_member(path, _SETTINGS_PART, new_settings)
    if windows_readonly:
        set_windows_readonly_flag(path, True)


def save_formed_protocol_docx(doc: object, output_path: str | Path) -> None:
    """
    Сохранить сформированный протокол и снять защиту Word, унаследованную от шаблона.
    """
    path = Path(output_path)
    doc.save(str(path))  # type: ignore[union-attr]
    remove_docx_word_readonly_protection(path, windows_readonly=False)


def remove_docx_word_readonly_protection(
    path: Path,
    *,
    windows_readonly: bool = True,
) -> None:
    path = Path(path)
    if windows_readonly:
        set_windows_readonly_flag(path, False)
    with zipfile.ZipFile(path, "r") as zf:
        settings = zf.read(_SETTINGS_PART)
    new_settings = _patch_settings_xml(settings, enable=False)
    if new_settings != settings:
        _rewrite_docx_member(path, _SETTINGS_PART, new_settings)
    elif windows_readonly:
        set_windows_readonly_flag(path, False)


def set_windows_readonly_flag(path: Path, readonly: bool) -> None:
    if sys.platform != "win32":
        return
    path = Path(path)
    if not path.is_file():
        return
    mode = path.stat().st_mode
    if readonly:
        path.chmod(mode & ~stat.S_IWRITE)
    else:
        path.chmod(mode | stat.S_IWUSR | stat.S_IWRITE)


def _all_template_paths(
    base_dir: Path | None,
    extra_paths: Iterable[Path | str] | None,
) -> list[Path]:
    extra: list[Path] = []
    if base_dir is not None:
        extra.extend(standard_protocol_template_paths(base_dir))
    if extra_paths:
        extra.extend(Path(p) for p in extra_paths)
    return discover_standard_protocol_template_paths(extra)


def protect_standard_protocol_templates(
    base_dir: Path | None = None,
    *,
    windows_readonly: bool = True,
    extra_paths: Iterable[Path | str] | None = None,
) -> list[str]:
    done: list[str] = []
    paths = _all_template_paths(base_dir, extra_paths)
    for p in paths:
        if is_docx_word_readonly_protected(p):
            if windows_readonly:
                set_windows_readonly_flag(p, True)
            done.append(f"(уже) {p}")
            continue
        apply_docx_word_readonly_protection(p, windows_readonly=windows_readonly)
        done.append(str(p))
    return done


def unprotect_standard_protocol_templates(
    base_dir: Path | None = None,
    *,
    extra_paths: Iterable[Path | str] | None = None,
) -> tuple[list[str], list[str]]:
    """
    Снять защиту со всех найденных шаблонов.

    Returns:
        (успешно, ошибки) — списки путей/сообщений.
    """
    paths = _all_template_paths(base_dir, extra_paths)
    ok: list[str] = []
    err: list[str] = []
    if not paths:
        return ok, err
    for p in paths:
        try:
            remove_docx_word_readonly_protection(p, windows_readonly=True)
            ok.append(str(p))
        except OSError as e:
            err.append(f"{p}\n  ({type(e).__name__}) {e}")
    return ok, err
