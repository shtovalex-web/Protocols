# -*- coding: utf-8 -*-
"""update_config.json — разбор путей Windows в JSON."""

from __future__ import annotations

import importlib.util
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
_NEXT = ROOT / "ProtocolOHT_next"
sys.path.insert(0, str(_NEXT))

from update_config import (  # noqa: E402
    UpdateConfig,
    format_manifest_path_for_json,
    parse_update_config_text,
    resolve_update_share_root,
)


class TestUpdateConfig(unittest.TestCase):
    def test_parse_valid_json_with_forward_slashes(self):
        cfg = parse_update_config_text(
            '{"manifest_path": "D:/Обновление/manifest.json", "enabled": true}'
        )
        self.assertIsNotNone(cfg)
        assert cfg is not None
        self.assertEqual(cfg.manifest_path, Path("D:/Обновление/manifest.json"))
        self.assertTrue(cfg.enabled)

    def test_parse_invalid_escape_windows_path_fallback(self):
        raw = '{\n  "manifest_path": "D:\\Обновление\\manifest.json",\n  "enabled": true\n}'
        cfg = parse_update_config_text(raw)
        self.assertIsNotNone(cfg)
        assert cfg is not None
        self.assertEqual(cfg.manifest_path, Path(r"D:\Обновление\manifest.json"))
        self.assertTrue(cfg.enabled)

    def test_format_manifest_path_for_json_uses_forward_slashes(self):
        self.assertEqual(
            format_manifest_path_for_json(Path(r"D:\Обновление\manifest.json")),
            "D:/Обновление/manifest.json",
        )

    def test_resolve_update_share_root_accepts_directory(self):
        self.assertEqual(
            resolve_update_share_root(Path(r"D:/Обновление")),
            Path(r"D:/Обновление"),
        )


if __name__ == "__main__":
    unittest.main()
