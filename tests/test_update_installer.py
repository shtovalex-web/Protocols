# -*- coding: utf-8
"""Установка обновления: копия и rename .exe."""

from __future__ import annotations

import hashlib
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from _bootstrap import setup_main_project_paths

setup_main_project_paths()

from update_installer import (  # noqa: E402
    cleanup_backup_exe,
    stage_payload_copy,
    swap_exe_via_rename,
)


class TestUpdateInstaller(unittest.TestCase):
    def test_stage_payload_copy_verifies_sha256(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source.exe"
            payload = b"payload-bytes"
            source.write_bytes(payload)
            digest = hashlib.sha256(payload).hexdigest()

            staged = stage_payload_copy(
                source,
                root / "app.exe.new",
                expected_sha256=digest,
                expected_size=len(payload),
            )
            self.assertEqual(staged.read_bytes(), payload)

    def test_swap_exe_via_rename(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            current = root / "app.exe"
            new_file = root / "app.exe.new"
            current.write_bytes(b"old")
            new_file.write_bytes(b"new")

            swap_exe_via_rename(current)

            self.assertEqual(current.read_bytes(), b"new")
            self.assertFalse(new_file.exists())
            backup = root / "app.exe.old"
            self.assertEqual(backup.read_bytes(), b"old")

            cleanup_backup_exe(current)
            self.assertFalse(backup.exists())

    def test_cleanup_backup_exe_ignores_permission_error(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            current = root / "app.exe"
            backup = root / "app.exe.old"
            backup.write_bytes(b"old")

            with patch.object(Path, "unlink", side_effect=PermissionError("locked")):
                self.assertFalse(cleanup_backup_exe(current))
            self.assertTrue(backup.is_file())

    def test_launch_updated_exe_uses_cmd_start_on_windows(self) -> None:
        from update_installer import launch_updated_exe

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            exe = root / "ProtocolOOT.exe"
            exe.write_bytes(b"exe")
            with patch("update_installer.sys.platform", "win32"):
                with patch("update_installer.subprocess.Popen") as popen:
                    launch_updated_exe(exe, show_changelog=True, version="1.6.1")
            popen.assert_called_once()
            cmd = popen.call_args.args[0]
            self.assertEqual(cmd[0:3], ["cmd", "/c", "start"])
            self.assertIn(str(exe.resolve()), cmd)
            self.assertIn("--show-changelog=1.6.1", cmd)


if __name__ == "__main__":
    unittest.main()
