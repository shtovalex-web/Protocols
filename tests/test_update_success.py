# -*- coding: utf-8
"""Уведомление об успешном обновлении."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from _bootstrap import setup_main_project_paths

setup_main_project_paths()

from update_installer import apply_pending_app_staging  # noqa: E402
from windows_app_bundle import APP_UPDATE_STAGING_DIR, INTERNAL_DIR_NAME  # noqa: E402


class TestApplyPendingStaging(unittest.TestCase):
    def test_apply_pending_app_staging_copies_files(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            staging = root / APP_UPDATE_STAGING_DIR
            internal = staging / INTERNAL_DIR_NAME
            internal.mkdir(parents=True)
            (staging / "ProtocolOOT.exe").write_bytes(b"new-exe")
            (internal / "python312.dll").write_bytes(b"dll")
            (root / "ProtocolOOT.exe").write_bytes(b"old-exe")

            self.assertTrue(apply_pending_app_staging(root))
            self.assertEqual((root / "ProtocolOOT.exe").read_bytes(), b"new-exe")
            self.assertEqual((root / INTERNAL_DIR_NAME / "python312.dll").read_bytes(), b"dll")
            self.assertFalse(staging.exists())

    def test_apply_pending_returns_false_when_empty(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            self.assertFalse(apply_pending_app_staging(Path(tmp)))


class TestUpdateSuccess(unittest.TestCase):
    def test_notify_shows_message_and_exits(self) -> None:
        from update_success import notify_update_success_and_exit

        with tempfile.TemporaryDirectory() as tmp:
            exe = Path(tmp) / "ProtocolOOT.exe"
            exe.write_bytes(b"x")
            with patch("update_success.messagebox.showinfo") as show:
                with patch("update_success.exit_after_update") as exit_fn:
                    notify_update_success_and_exit(
                        version="1.6.2",
                        exe_path=exe,
                        parent=None,
                        bundle_staged=False,
                    )
            show.assert_called_once()
            self.assertIn("1.6.2", show.call_args.args[1])
            self.assertIn("ProtocolOOT.exe", show.call_args.args[1])
            exit_fn.assert_called_once()


if __name__ == "__main__":
    unittest.main()
