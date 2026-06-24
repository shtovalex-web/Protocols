# -*- coding: utf-8
"""Файлы комплекта data/ для автообновления (политика replace)."""

from __future__ import annotations

from pathlib import Path
# Только каталог data/ рядом с exe. Файлы в корне (protocols.db, Data_base.xlsx…) не обновляются.
DATA_REPLACE_FILENAMES: tuple[str, ...] = (
    "default_protocol.docx",
    "default_protocol_tehnicheskiy.docx",
    "ПОДРОБНАЯ_ИНСТРУКЦИЯ_для_пользователя.docx",
    "ИНСТРУКЦИЯ_оформление_протоколов_Минтруд.docx",
    "ЖУРНАЛ_ДОРАБОТОК.md",
    "FAQ.txt",
    "Шаблон_Минтруд_XSD_УМН.xlsx",
    "!! Шаблон_Минтруд_XSD_УМН _ общ+.xlsx",
    "Шаблон_Минтруд_XSD_УМН _ общ+.xlsx",
    "icon.ico",
)

DATA_SUBDIR_NAME = "data"
DATA_POLICY_REPLACE = "replace"


def build_data_manifest_entries(
    *,
    data_src_dir: Path,
    version: str | None = None,
    paths_relative_to_version_dir: bool = False,
) -> list[dict[str, object]]:
    from update_manifest import sha256_file

    entries: list[dict[str, object]] = []
    for name in DATA_REPLACE_FILENAMES:
        src = data_src_dir / name
        if not src.is_file():
            continue
        if paths_relative_to_version_dir:
            rel = f"{DATA_SUBDIR_NAME}/{name}".replace("\\", "/")
        else:
            if not version:
                msg = "version is required when paths_relative_to_version_dir is False"
                raise ValueError(msg)
            rel = f"windows/{version}/{DATA_SUBDIR_NAME}/{name}".replace("\\", "/")
        entries.append(
            {
                "relative_path": rel,
                "sha256": sha256_file(src),
                "size": src.stat().st_size,
                "policy": DATA_POLICY_REPLACE,
            }
        )
    return entries
