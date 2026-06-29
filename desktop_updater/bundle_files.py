# -*- coding: utf-8
"""Файлы комплекта data/ для автообновления (политика replace)."""

from __future__ import annotations

from pathlib import Path

from desktop_updater.registry import get_config

DATA_POLICY_REPLACE = "replace"


def data_subdir_name() -> str:
    return get_config().data_subdir


def data_replace_filenames() -> tuple[str, ...]:
    return get_config().data_replace_filenames


def build_data_manifest_entries(
    *,
    data_src_dir: Path,
    version: str | None = None,
    paths_relative_to_version_dir: bool = False,
) -> list[dict[str, object]]:
    from desktop_updater.manifest import sha256_file

    cfg = get_config()
    subdir = cfg.data_subdir
    entries: list[dict[str, object]] = []
    for name in cfg.data_replace_filenames:
        src = data_src_dir / name
        if not src.is_file():
            continue
        if paths_relative_to_version_dir:
            rel = f"{subdir}/{name}".replace("\\", "/")
        else:
            if not version:
                msg = "version is required when paths_relative_to_version_dir is False"
                raise ValueError(msg)
            rel = f"windows/{version}/{subdir}/{name}".replace("\\", "/")
        entries.append(
            {
                "relative_path": rel,
                "sha256": sha256_file(src),
                "size": src.stat().st_size,
                "policy": DATA_POLICY_REPLACE,
            }
        )
    return entries
