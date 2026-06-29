# -*- coding: utf-8 -*-
"""data/update_info.json — маркер версии комплекта."""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
NEXT = ROOT / "ProtocolOHT_next"
sys.path.insert(0, str(NEXT))

from update_info import (  # noqa: E402
    UPDATE_INFO_FILENAME,
    installed_version_from_data,
    load_update_info,
    write_update_info,
)


class TestUpdateInfo(unittest.TestCase):
    def test_write_and_load_update_info(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            data_dir = Path(tmp)
            path = write_update_info(
                data_dir,
                version="1.5.3",
                released="2026-06-09",
                platform="windows",
            )
            self.assertEqual(path.name, UPDATE_INFO_FILENAME)
            payload = json.loads(path.read_text(encoding="utf-8"))
            self.assertEqual(payload["version"], "1.5.3")
            self.assertEqual(payload["released"], "2026-06-09")
            self.assertEqual(payload["platform"], "windows")
            info = load_update_info(data_dir)
            assert info is not None
            self.assertEqual(info.version, "1.5.3")

    def test_installed_version_from_data(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            data_dir = root / "data"
            data_dir.mkdir()
            write_update_info(data_dir, version="1.5.2", released="2026-06-01")
            exe = root / "ProtocolOOT.exe"
            exe.write_bytes(b"stub")
            self.assertEqual(installed_version_from_data(exe), "1.5.2")

    def test_installed_version_missing_returns_none(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            exe = root / "ProtocolOOT.exe"
            exe.write_bytes(b"stub")
            self.assertIsNone(installed_version_from_data(exe))


if __name__ == "__main__":
    unittest.main()
