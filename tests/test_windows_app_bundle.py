# -*- coding: utf-8
"""Windows onedir bundle zip."""

from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path

from _bootstrap import setup_main_project_paths

setup_main_project_paths()

from windows_app_bundle import (  # noqa: E402
    APP_UPDATE_STAGING_DIR,
    WINDOWS_APP_ZIP_NAME,
    create_windows_app_zip,
    is_windows_app_bundle,
)


class TestWindowsAppBundle(unittest.TestCase):
    def test_create_and_detect_zip(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            exe = root / "ProtocolOOT.exe"
            internal = root / "_internal"
            internal.mkdir()
            exe.write_bytes(b"exe")
            (internal / "python313.dll").write_bytes(b"dll")
            zip_path = create_windows_app_zip(root)
            self.assertTrue(is_windows_app_bundle(zip_path))
            self.assertEqual(zip_path.name, WINDOWS_APP_ZIP_NAME)
            with zipfile.ZipFile(zip_path) as archive:
                names = set(archive.namelist())
            self.assertIn("ProtocolOOT.exe", names)
            self.assertIn("_internal/python313.dll", names)

    def test_staging_dir_name(self) -> None:
        self.assertEqual(APP_UPDATE_STAGING_DIR, ".app_update_staging")


if __name__ == "__main__":
    unittest.main()
