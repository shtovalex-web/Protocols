# -*- coding: utf-8 -*-
"""Реестр обученных Минтруда: проверка доступности пути."""

from __future__ import annotations

import os
import tempfile
import unittest

from _bootstrap import setup_main_project_paths

setup_main_project_paths()

from mintrud_trained_registry import (
    check_trained_registry_file_access,
    normalize_trained_registry_path,
    trained_registry_status_user_message,
)


class TestMintrudTrainedRegistryPath(unittest.TestCase):
    def test_normalize_empty(self) -> None:
        self.assertIsNone(normalize_trained_registry_path(""))
        self.assertIsNone(normalize_trained_registry_path(None))

    def test_check_not_set(self) -> None:
        self.assertEqual(check_trained_registry_file_access(""), "not_set")
        self.assertEqual(check_trained_registry_file_access(None), "not_set")

    def test_check_ok_existing_file(self) -> None:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name
        try:
            self.assertEqual(check_trained_registry_file_access(path), "ok")
        finally:
            os.unlink(path)

    def test_check_not_found(self) -> None:
        missing = os.path.join(tempfile.gettempdir(), "protocoloot_missing_registry_test.xlsx")
        self.assertEqual(check_trained_registry_file_access(missing), "not_found")

    def test_status_messages(self) -> None:
        self.assertIn(
            "не выбран",
            trained_registry_status_user_message("not_set", "").lower(),
        )
        self.assertIn(
            "используется",
            trained_registry_status_user_message("ok", r"C:\tmp\a.xlsx").lower(),
        )
        self.assertIn(
            "не найден",
            trained_registry_status_user_message("not_found", r"C:\tmp\a.xlsx").lower(),
        )
        self.assertIn(
            "нет доступа",
            trained_registry_status_user_message("inaccessible", r"\\srv\a.xlsx").lower(),
        )


if __name__ == "__main__":
    unittest.main()
