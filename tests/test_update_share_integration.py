# -*- coding: utf-8
"""Интеграция: манифест D:\\Обновление и логика предложения обновления."""

from __future__ import annotations

import json
import unittest
from pathlib import Path
from unittest.mock import patch

from _bootstrap import setup_main_project_paths

setup_main_project_paths()

from startup_update import prepare_startup_updates  # noqa: E402

SHARE_MANIFEST = Path(r"D:\Обновление\manifest.json")
EXE_DIR = Path(__file__).resolve().parents[1] / "ProtocolOHT_onefile"
EXE_PATH = EXE_DIR / "ProtocolOOT.exe"


@unittest.skipUnless(SHARE_MANIFEST.is_file(), "нет D:\\Обновление\\manifest.json")
@unittest.skipUnless(EXE_PATH.is_file(), "нет собранного ProtocolOOT.exe")
class TestUpdateShareIntegration(unittest.TestCase):
    def test_frozen_startup_offers_update_then_declines(self) -> None:
        config = {
            "manifest_path": str(SHARE_MANIFEST),
            "enabled": True,
        }
        config_path = EXE_DIR / "update_config.json"
        config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")

        asked: list[bool] = []

        def _ask_yesno(title, text, parent=None):
            asked.append(True)
            self.assertIn("Доступна новая версия", text)
            return False

        with patch("startup_update.is_frozen", return_value=True):
            with patch("startup_update.current_exe_path", return_value=EXE_PATH):
                with patch("startup_update.app_version", return_value="1.6.0"):
                    with patch("startup_update.load_update_config") as load_cfg:
                        from update_config import UpdateConfig

                        load_cfg.return_value = UpdateConfig(manifest_path=SHARE_MANIFEST)
                        with patch("startup_update.messagebox.askyesno", side_effect=_ask_yesno):
                            self.assertTrue(prepare_startup_updates([str(EXE_PATH)]))
        self.assertTrue(asked, "диалог обновления не был показан")


if __name__ == "__main__":
    unittest.main()
