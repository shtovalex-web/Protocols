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

    def test_launch_updated_exe_uses_cmd_helper_on_windows(self) -> None:
        from update_installer import launch_updated_exe

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            exe = root / "ProtocolOOT.exe"
            exe.write_bytes(b"exe")
            with patch("desktop_updater.installer.sys.platform", "win32"):
                with patch(
                    "desktop_updater.installer._launch_updated_exe_cmd_helper",
                    return_value=True,
                ) as helper:
                    launch_updated_exe(exe)
            helper.assert_called_once()

    def test_restart_cmd_uses_dp0_and_disable_delayed_expansion(self) -> None:
        from update_installer import _write_restart_cmd

        with tempfile.TemporaryDirectory(prefix="test path_") as tmp:
            root = Path(tmp) / "!Протоколы"
            root.mkdir(parents=True)
            exe = root / "ProtocolOOT.exe"
            exe.write_bytes(b"exe")
            cmd_path = _write_restart_cmd(exe, parent_pid=12345)
            text = cmd_path.read_text(encoding="ascii")
            self.assertIn("DisableDelayedExpansion", text)
            self.assertIn("set TCL_LIBRARY=", text)
            self.assertIn("PID eq 12345", text)
            self.assertIn(":wait_parent", text)
            self.assertIn("robocopy", text)
            self.assertIn(".app_update_staging", text)
            self.assertIn('start "" /D "%~dp0" "%~dp0ProtocolOOT.exe"', text)
            self.assertNotIn("Протоколы", text)

    def test_launch_updated_exe_falls_back_to_shell_execute(self) -> None:
        from update_installer import launch_updated_exe

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            exe = root / "ProtocolOOT.exe"
            exe.write_bytes(b"exe")
            with patch("desktop_updater.installer.sys.platform", "win32"):
                with patch("desktop_updater.installer._launch_updated_exe_cmd_helper", return_value=False):
                    with patch("desktop_updater.installer._launch_updated_exe_windows", return_value=True) as shell:
                        launch_updated_exe(exe)
            shell.assert_called_once()

    def test_launch_updated_exe_falls_back_to_subprocess(self) -> None:
        from update_installer import launch_updated_exe

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            exe = root / "ProtocolOOT.exe"
            exe.write_bytes(b"exe")
            with patch("desktop_updater.installer.sys.platform", "win32"):
                with patch("desktop_updater.installer._launch_updated_exe_cmd_helper", return_value=False):
                    with patch("desktop_updater.installer._launch_updated_exe_windows", return_value=False):
                        with patch("desktop_updater.installer._launch_updated_exe_subprocess") as sub:
                            launch_updated_exe(exe)
            sub.assert_called_once()

    def test_launch_updated_exe_handles_path_with_spaces_and_bang(self) -> None:
        from update_installer import _launch_updated_exe_subprocess

        with tempfile.TemporaryDirectory(prefix="test path_") as tmp:
            root = Path(tmp) / "!folder name"
            root.mkdir(parents=True)
            exe = root / "ProtocolOOT.exe"
            exe.write_bytes(b"exe")
            with patch("desktop_updater.installer.subprocess.Popen") as popen:
                _launch_updated_exe_subprocess(exe, cwd=str(root))
            cmd = popen.call_args.args[0]
            self.assertEqual(cmd[0], str(exe))


if __name__ == "__main__":
    unittest.main()
