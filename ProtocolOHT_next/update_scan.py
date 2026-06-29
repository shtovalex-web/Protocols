# -*- coding: utf-8
"""Поиск обновлений в каталогах платформ шары: windows/<версия>/, linux/<версия>/."""

from __future__ import annotations

import sys
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
DEFAULT_LINUX_EXE_NAME = "ProtocolOOT"
UPDATE_PLATFORM_SUBDIRS = ("windows", "linux")
PLATFORM_EXE_NAMES = {
    "windows": DEFAULT_WINDOWS_EXE_NAME,
    "linux": DEFAULT_LINUX_EXE_NAME,
}


@dataclass(frozen=True)
class UpdateCandidate:
    version: str
    manifest: UpdateManifest
    anchor_manifest_path: Path
    platform: str
    source: str


def default_update_platform() -> str:
    return "windows" if sys.platform == "win32" else "linux"


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
    exe_path: Path,
    version: str,
    changes: list[str] | None = None,
) -> UpdateManifest:
    size = exe_path.stat().st_size
    digest = sha256_file(exe_path)
    return UpdateManifest(
        latest_version=version,
        windows=WindowsUpdatePayload(
            relative_path=exe_path.name,
            sha256=digest,
            size=size,
        ),
        changes_short=changes or [f"Обновление до версии {version}"],
    )


def _candidate_from_manifest_file(manifest_path: Path, *, platform: str) -> UpdateCandidate | None:
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
        platform=platform,
        source=f"manifest:{manifest_path.parent.name}",
    )


def _candidate_from_version_dir(
    *,
    version_dir: Path,
    platform: str,
    exe_name: str,
) -> UpdateCandidate | None:
    version = version_from_dir_name(version_dir.name)
    if version is None:
        return None

    manifest_path = version_dir / "manifest.json"
    if manifest_path.is_file():
        return _candidate_from_manifest_file(manifest_path, platform=platform)

    exe_path = version_dir / exe_name
    if not exe_path.is_file():
        return None

    manifest = manifest_from_exe(exe_path=exe_path, version=version)
    return UpdateCandidate(
        version=version,
        manifest=manifest,
        anchor_manifest_path=manifest_path,
        platform=platform,
        source=f"exe:{platform}/{version_dir.name}/{exe_name}",
    )


def _scan_platform_dir(
    share_root: Path,
    *,
    platform: str,
    exe_name: str | None = None,
) -> list[UpdateCandidate]:
    platform_root = share_root / platform
    if not platform_root.is_dir():
        return []

    payload_name = exe_name or PLATFORM_EXE_NAMES.get(platform, DEFAULT_WINDOWS_EXE_NAME)
    found: dict[str, UpdateCandidate] = {}

    for entry in sorted(platform_root.iterdir(), key=lambda item: item.name):
        if not entry.is_dir():
            continue
        candidate = _candidate_from_version_dir(
            version_dir=entry,
            platform=platform,
            exe_name=payload_name,
        )
        if candidate is None:
            continue
        found[candidate.version] = candidate

    return list(found.values())


def scan_update_candidates(
    share_root: Path,
    *,
    platform: str | None = None,
    exe_name: str | None = None,
) -> list[UpdateCandidate]:
    """Сканирует windows/<версия>/ и linux/<версия>/ (или одну платформу)."""
    root = share_root.expanduser().resolve()
    if not root.is_dir():
        return []

    platforms = (platform,) if platform else UPDATE_PLATFORM_SUBDIRS
    found: dict[tuple[str, str], UpdateCandidate] = {}

    for platform_name in platforms:
        for candidate in _scan_platform_dir(root, platform=platform_name, exe_name=exe_name):
            found[(candidate.platform, candidate.version)] = candidate

    return sorted(found.values(), key=lambda item: (item.platform, parse_version(item.version)))


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
    platform: str | None = None,
    exe_name: str | None = None,
) -> UpdateCandidate | None:
    """Сканирование каталогов платформ; выбор новейшей версии выше текущей."""
    root = share_root.expanduser().resolve()
    if not root.is_dir():
        return None
    target_platform = platform if platform is not None else default_update_platform()
    candidates = scan_update_candidates(root, platform=target_platform, exe_name=exe_name)
    return pick_newest_update(candidates, current_version)
