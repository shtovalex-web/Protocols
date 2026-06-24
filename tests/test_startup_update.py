# -*- coding: utf-8
"""Проверка обновлений при старте."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from _bootstrap import setup_main_project_paths

setup_main_project_paths()

from startup_update import prepare_startup_updates  # noqa: E402


class TestStartupUpdate(unittest.TestCase):
    def test_prepare_startup_updates_skips_when_not_frozen(self) -> None:
        with patch("startup_update.is_frozen", return_value=False):
            with patch.dict("os.environ", {}, clear=False):
                import os

                os.environ.pop("PROTOCOLOOT_UPDATE_CHECK", None)
                self.assertTrue(prepare_startup_updates(["app"]))

    def test_prepare_startup_updates_no_manifest_is_silent(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            missing = Path(tmp) / "missing.json"
            env = {
                "PROTOCOLOOT_UPDATE_CHECK": "1",
                "PROTOCOLOOT_UPDATE_MANIFEST": str(missing),
            }
            with patch("startup_update.is_frozen", return_value=True):
                with patch.dict("os.environ", env, clear=False):
                    self.assertTrue(prepare_startup_updates(["app"]))


if __name__ == "__main__":
    unittest.main()
