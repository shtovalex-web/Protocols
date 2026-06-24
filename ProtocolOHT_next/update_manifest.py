# -*- coding: utf-8 -*-
"""Чтение manifest.json с сетевой шары."""

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass, field
from pathlib import Path


from update_bundle_files import DATA_POLICY_REPLACE


class UpdateManifestError(Exception):
    """Ошибка чтения или проверки манифеста."""


@dataclass
class WindowsUpdatePayload:
    relative_path: str
    sha256: str
    size: int

    @classmethod
    def from_dict(cls, data: object) -> WindowsUpdatePayload:
        if not isinstance(data, dict):
            msg = "manifest.windows must be an object"
            raise UpdateManifestError(msg)
        rel = str(data.get("relative_path", "")).strip()
        sha = str(data.get("sha256", "")).strip()
        try:
            size = int(data.get("size", 0))
        except (TypeError, ValueError) as error:
            msg = "manifest.windows.size must be a positive integer"
            raise UpdateManifestError(msg) from error
        if not rel or not sha or size < 1:
            msg = "manifest.windows is incomplete"
            raise UpdateManifestError(msg)
        return cls(relative_path=rel, sha256=sha, size=size)


@dataclass
class DataFilePayload:
    relative_path: str
    sha256: str
    size: int
    policy: str = DATA_POLICY_REPLACE

    @classmethod
    def from_dict(cls, data: object) -> DataFilePayload:
        if not isinstance(data, dict):
            msg = "manifest.data_files[] entry must be an object"
            raise UpdateManifestError(msg)
        rel = str(data.get("relative_path", "")).strip()
        sha = str(data.get("sha256", "")).strip()
        policy = str(data.get("policy", DATA_POLICY_REPLACE)).strip() or DATA_POLICY_REPLACE
        try:
            size = int(data.get("size", 0))
        except (TypeError, ValueError) as error:
            msg = "manifest.data_files[].size must be a positive integer"
            raise UpdateManifestError(msg) from error
        if not rel or not sha or size < 1:
            msg = "manifest.data_files[] entry is incomplete"
            raise UpdateManifestError(msg)
        return cls(relative_path=rel, sha256=sha, size=size, policy=policy)


@dataclass
class UpdateManifest:
    latest_version: str
    windows: WindowsUpdatePayload
    released: str = ""
    mandatory: bool = False
    changes_short: list[str] = field(default_factory=list)
    data_files: list[DataFilePayload] = field(default_factory=list)

    def windows_payload_path(self, manifest_path: Path) -> Path:
        return manifest_path.parent / self.windows.relative_path

    def replace_data_files(self) -> list[DataFilePayload]:
        return [item for item in self.data_files if item.policy == DATA_POLICY_REPLACE]

    @classmethod
    def from_dict(cls, data: object) -> UpdateManifest:
        if not isinstance(data, dict):
            msg = "manifest root must be an object"
            raise UpdateManifestError(msg)
        version = str(data.get("latest_version", "")).strip()
        if not version:
            msg = "manifest.latest_version is required"
            raise UpdateManifestError(msg)
        windows = WindowsUpdatePayload.from_dict(data.get("windows"))
        released = str(data.get("released", "") or "")
        mandatory = bool(data.get("mandatory", False))
        raw_changes = data.get("changes_short", [])
        changes: list[str] = []
        if isinstance(raw_changes, list):
            changes = [str(item) for item in raw_changes if str(item).strip()]
        data_files: list[DataFilePayload] = []
        raw_data = data.get("data_files", [])
        if isinstance(raw_data, list):
            for item in raw_data:
                data_files.append(DataFilePayload.from_dict(item))
        return cls(
            latest_version=version,
            windows=windows,
            released=released,
            mandatory=mandatory,
            changes_short=changes,
            data_files=data_files,
        )


def load_update_manifest(manifest_path: Path) -> UpdateManifest:
    if not manifest_path.is_file():
        msg = f"Manifest not found: {manifest_path}"
        raise UpdateManifestError(msg)
    try:
        data = json.loads(manifest_path.read_text(encoding="utf-8"))
        return UpdateManifest.from_dict(data)
    except (OSError, json.JSONDecodeError, UpdateManifestError) as error:
        msg = f"Failed to read manifest: {manifest_path}"
        raise UpdateManifestError(msg) from error


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()
