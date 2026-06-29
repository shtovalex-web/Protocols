# -*- coding: utf-8
"""Сканирование шары обновлений."""

from __future__ import annotations

import hashlib
import json
import tempfile
import unittest
from pathlib import Path

from _bootstrap import setup_main_project_paths

setup_main_project_paths()

from update_scan import (  # noqa: E402
    resolve_latest_update,
    scan_update_candidates,
)


class TestUpdateScan(unittest.TestCase):
    def _write_exe(self, path: Path, payload: bytes = b"exe-payload") -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(payload)

    def test_scan_version_subfolders_under_windows(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            self._write_exe(root / "windows" / "1.6.0" / "ProtocolOOT.exe", b"v160")
            self._write_exe(root / "windows" / "1.6.2" / "ProtocolOOT.exe", b"v162")

            candidates = scan_update_candidates(root, platform="windows")
            versions = [c.version for c in candidates]
            self.assertEqual(versions, ["1.6.0", "1.6.2"])
            self.assertTrue(all(c.platform == "windows" for c in candidates))

    def test_scan_skips_non_version_dirs_and_root_manifest(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            self._write_exe(root / "windows" / "1.6.0" / "ProtocolOOT.exe", b"v160")
            (root / "windows" / "draft").mkdir()
            (root / "windows" / "draft" / "ProtocolOOT.exe").write_bytes(b"skip")
            digest = hashlib.sha256(b"legacy").hexdigest()
            (root / "manifest.json").write_text(
                json.dumps(
                    {
                        "latest_version": "9.9.9",
                        "windows": {
                            "relative_path": "windows/9.9.9/ProtocolOOT.exe",
                            "sha256": digest,
                            "size": len(b"legacy"),
                        },
                    }
                ),
                encoding="utf-8",
            )

            candidates = scan_update_candidates(root, platform="windows")
            self.assertEqual([c.version for c in candidates], ["1.6.0"])

    def test_scan_linux_subdir_when_platform_none(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            self._write_exe(root / "linux" / "2.0.0" / "ProtocolOOT", b"linux")

            candidates = scan_update_candidates(root)
            self.assertEqual(len(candidates), 1)
            self.assertEqual(candidates[0].platform, "linux")
            self.assertEqual(candidates[0].version, "2.0.0")

    def test_resolve_picks_newest_newer_than_current(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            exe_old = root / "windows" / "1.6.0" / "ProtocolOOT.exe"
            exe_new = root / "windows" / "1.6.2" / "ProtocolOOT.exe"
            self._write_exe(exe_old, b"old")
            self._write_exe(exe_new, b"new")

            resolved = resolve_latest_update(root, current_version="1.6.0", platform="windows")
            self.assertIsNotNone(resolved)
            assert resolved is not None
            self.assertEqual(resolved.version, "1.6.2")
            self.assertEqual(resolved.manifest.windows.size, len(b"new"))

    def test_resolve_returns_none_when_up_to_date(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            exe_path = root / "windows" / "1.6.0" / "ProtocolOOT.exe"
            self._write_exe(exe_path, b"v160")

            self.assertIsNone(
                resolve_latest_update(root, current_version="1.6.0", platform="windows")
            )

    def test_nested_manifest_json(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            nested = root / "windows" / "1.7.0"
            exe_path = nested / "ProtocolOOT.exe"
            self._write_exe(exe_path, b"v170")
            digest = hashlib.sha256(b"v170").hexdigest()
            (nested / "manifest.json").write_text(
                json.dumps(
                    {
                        "latest_version": "1.7.0",
                        "changes_short": ["Релиз 1.7"],
                        "windows": {
                            "relative_path": "ProtocolOOT.exe",
                            "sha256": digest,
                            "size": len(b"v170"),
                        },
                    }
                ),
                encoding="utf-8",
            )

            resolved = resolve_latest_update(root, current_version="1.6.0", platform="windows")
            self.assertIsNotNone(resolved)
            assert resolved is not None
            self.assertEqual(resolved.version, "1.7.0")
            self.assertEqual(resolved.manifest.changes_short, ["Релиз 1.7"])

    def test_share_root_from_manifest(self) -> None:
        from update_config import resolve_update_share_root

        self.assertEqual(
            resolve_update_share_root(Path(r"D:\Обновление\manifest.json")),
            Path(r"D:\Обновление"),
        )
        self.assertEqual(
            resolve_update_share_root(Path(r"D:\Обновление")),
            Path(r"D:\Обновление"),
        )
        self.assertEqual(
            resolve_update_share_root(Path(r"D:\Обновление\windows\1.5.2\manifest.json")),
            Path(r"D:\Обновление"),
        )


if __name__ == "__main__":
    unittest.main()
