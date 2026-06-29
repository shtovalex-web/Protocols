# -*- coding: utf-8
"""Проверка manifest.json после публикации."""

from __future__ import annotations

import hashlib
import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "ProtocolOHT_next"))

from publish_update_manifest import refresh_manifest, verify_manifest_payload  # noqa: E402


class TestPublishManifestVerify(unittest.TestCase):
    def test_verify_manifest_payload_detects_mismatch(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            version_dir = Path(tmp) / "windows" / "1.0.0"
            data_dir = version_dir / "data"
            data_dir.mkdir(parents=True)
            payload = b"template"
            (data_dir / "FAQ.txt").write_bytes(payload)
            digest = hashlib.sha256(payload).hexdigest()
            exe = version_dir / "ProtocolOOT.exe"
            exe.write_bytes(b"exe")
            manifest_path = version_dir / "manifest.json"
            manifest_path.write_text(
                json.dumps(
                    {
                        "latest_version": "1.0.0",
                        "windows": {
                            "relative_path": "ProtocolOOT.exe",
                            "sha256": hashlib.sha256(b"exe").hexdigest(),
                            "size": 3,
                        },
                        "data_files": [
                            {
                                "relative_path": "data/FAQ.txt",
                                "sha256": "00" * 32,
                                "size": len(payload),
                                "policy": "replace",
                            }
                        ],
                    }
                ),
                encoding="utf-8",
            )
            errors = verify_manifest_payload(manifest_path)
            self.assertTrue(errors)

    def test_refresh_manifest_matches_files(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            version_dir = Path(tmp) / "windows" / "1.0.1"
            data_dir = version_dir / "data"
            data_dir.mkdir(parents=True)
            (data_dir / "FAQ.txt").write_text("faq", encoding="utf-8")
            (version_dir / "ProtocolOOT.exe").write_bytes(b"exe")

            manifest_path = refresh_manifest(
                version_dir=version_dir,
                version="1.0.1",
                changes=["test"],
                mandatory=False,
                released="2026-06-29",
            )
            self.assertFalse(verify_manifest_payload(manifest_path))


if __name__ == "__main__":
    unittest.main()
