# -*- coding: utf-8 -*-
"""Тесты определения файлов, влияющих на Linux-копию."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
LINUX_PORT = ROOT / "linux_port"
sys.path.insert(0, str(LINUX_PORT))

from sync_util import path_affects_linux_app  # noqa: E402


class TestLinuxSyncPaths(unittest.TestCase):
    def test_protocol_ui_affects_linux(self) -> None:
        self.assertTrue(path_affects_linux_app("ProtocolOHT_next/protocol_ui.py"))

    def test_ui_theme_affects_linux(self) -> None:
        self.assertTrue(path_affects_linux_app("ProtocolOHT_next/ui_theme.py"))

    def test_commission_admin_affects_linux(self) -> None:
        self.assertTrue(path_affects_linux_app("commission_admin.py"))

    def test_overlay_affects_linux(self) -> None:
        self.assertTrue(path_affects_linux_app("linux_port/overlays/protocol_output.py"))

    def test_tests_do_not_affect_linux(self) -> None:
        self.assertFalse(path_affects_linux_app("tests/test_ui_theme.py"))

    def test_linux_app_copy_ignored(self) -> None:
        self.assertFalse(path_affects_linux_app("linux_port/app/main.py"))


if __name__ == "__main__":
    unittest.main()
