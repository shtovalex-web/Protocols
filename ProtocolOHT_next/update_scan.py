# -*- coding: utf-8
"""Поиск обновлений во вложенных каталогах шары (manifest.json и ProtocolOOT.exe)."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from update_manifest import (
    UpdateManifest,
    UpdateManifestError,
    WindowsUpdatePayload,
    load_update_manifest,
    sha256_file,
)
from version_compare import is_newer_version, parse_version

DEFAULT_WINDOWS_EXE_NAME = "ProtocolOOT.exe"
_WINDOWS_SUBDIR = "windows"


@dataclass(frozen=True)
class UpdateCandidate:
    version: str
    manifest: UpdateManifest
    anchor_manifest_path: Path
    source: str


def share_root_from_manifest(manifest_path: Path) -> Path:
    from update_config import resolve_update_share_root

    return resolve_update_share_root(manifest_path)


def version_from_dir_name(name: str) -> str | None:
    try:
        parse_version(name)
    except ValueError:
        return None
    return name.strip()


def manifest_from_exe(
    *,
    share_root: Path,
    exe_path: Path,
    version: str,
    changes: list[str] | None = None,
) -> UpdateManifest:
    resolved_exe = exe_path.resolve()
    rel = resolved_exe.relative_to(share_root.resolve())
    size = resolved_exe.stat().st_size
    digest = sha256_file(resolved_exe)
    return UpdateManifest(
        latest_version=version,
        windows=WindowsUpdatePayload(
            relative_path=str(rel).replace("\\", "/"),
            sha256=digest,
            size=size,
        ),
        changes_short=changes or [f"Обновление до версии {version}"],
    )


def _candidate_from_manifest_file(manifest_path: Path) -> UpdateCandidate | None:
    try:
        manifest = load_update_manifest(manifest_path)
    except (UpdateManifestError, OSError):
        return None
    payload = manifest.windows_payload_path(manifest_path)
    if not payload.is_file():
        return None
    return UpdateCandidate(
        version=manifest.latest_version,
        manifest=manifest,
        anchor_manifest_path=manifest_path,
        source=f"manifest:{manifest_path.name}",
    )


def _candidate_from_version_exe(
    *,
    share_root: Path,
    exe_path: Path,
    version: str,
) -> UpdateCandidate | None:
    if not exe_path.is_file():
        return None
    version_dir = exe_path.parent
    manifest_file = version_dir / "manifest.json"
    if manifest_file.is_file():
        return _candidate_from_manifest_file(manifest_file)
    manifest = manifest_from_exe(share_root=share_root, exe_path=exe_path, version=version)
    return UpdateCandidate(
        version=version,
        manifest=manifest,
        anchor_manifest_path=share_root / "manifest.json",
        source=f"exe:{exe_path.relative_to(share_root)}",
    )


def scan_update_candidates(
    share_root: Path,
    *,
    exe_name: str = DEFAULT_WINDOWS_EXE_NAME,
) -> list[UpdateCandidate]:
    """Сканирует share_root: manifest.json во вложенных папках и exe в каталогах версий."""
    root = share_root.expanduser().resolve()
    if not root.is_dir():
        return []

    found: dict[str, UpdateCandidate] = {}

    for manifest_path in root.rglob("manifest.json"):
        candidate = _candidate_from_manifest_file(manifest_path)
        if candidate is None:
            continue
        found[candidate.version] = candidate

    for exe_path in root.rglob(exe_name):
        version = version_from_dir_name(exe_path.parent.name)
        if version is None:
            continue
        candidate = _candidate_from_version_exe(
            share_root=root,
            exe_path=exe_path,
            version=version,
        )
        if candidate is None:
            continue
        existing = found.get(version)
        if existing is None or existing.source.startswith("exe:"):
            found[version] = candidate

    return sorted(found.values(), key=lambda item: parse_version(item.version))


def pick_newest_update(
    candidates: list[UpdateCandidate],
    current_version: str,
) -> UpdateCandidate | None:
    newer = [c for c in candidates if is_newer_version(c.version, current_version)]
    if not newer:
        return None
    return max(newer, key=lambda item: parse_version(item.version))


def resolve_latest_update(
    share_root: Path,
    *,
    current_version: str,
    exe_name: str = DEFAULT_WINDOWS_EXE_NAME,
) -> UpdateCandidate | None:
    """Сканирование каталога шары; выбор новейшей версии выше текущей."""
    root = share_root.expanduser().resolve()
    if not root.is_dir():
        return None
    candidates = scan_update_candidates(root, exe_name=exe_name)
    return pick_newest_update(candidates, current_version)
