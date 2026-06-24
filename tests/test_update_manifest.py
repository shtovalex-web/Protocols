# -*- coding: utf-8
"""Чтение manifest.json."""

from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from _bootstrap import setup_main_project_paths

setup_main_project_paths()

from update_manifest import UpdateManifestError, load_update_manifest  # noqa: E402


class TestUpdateManifest(unittest.TestCase):
    def test_load_update_manifest(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            manifest_path = root / "manifest.json"
            manifest_path.write_text(
                json.dumps(
                    {
                        "latest_version": "1.6.0",
                        "windows": {
                            "relative_path": "windows/1.6.0/ProtocolOOT.exe",
                            "sha256": "ab" * 32,
                            "size": 12345,
                        },
                        "changes_short": ["Тест"],
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            manifest = load_update_manifest(manifest_path)
            self.assertEqual(manifest.latest_version, "1.6.0")
            self.assertEqual(manifest.changes_short, ["Тест"])
            self.assertEqual(
                manifest.windows_payload_path(manifest_path),
                root / "windows" / "1.6.0" / "ProtocolOOT.exe",
            )

    def test_load_update_manifest_with_data_files(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            manifest_path = root / "manifest.json"
            manifest_path.write_text(
                json.dumps(
                    {
                        "latest_version": "1.5.2",
                        "windows": {
                            "relative_path": "ProtocolOOT.exe",
                            "sha256": "ab" * 32,
                            "size": 100,
                        },
                        "data_files": [
                            {
                                "relative_path": "data/FAQ.txt",
                                "sha256": "cd" * 32,
                                "size": 50,
                                "policy": "replace",
                            }
                        ],
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            manifest = load_update_manifest(manifest_path)
            self.assertEqual(len(manifest.replace_data_files()), 1)
            self.assertEqual(manifest.data_files[0].relative_path, "data/FAQ.txt")

    def test_missing_manifest_raises(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaises(UpdateManifestError):
                load_update_manifest(Path(tmp) / "missing.json")


if __name__ == "__main__":
    unittest.main()
