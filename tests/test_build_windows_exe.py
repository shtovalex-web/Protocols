# -*- coding: utf-8 -*-
"""Комплект Windows-сборки (build_windows_exe.py)."""

from __future__ import annotations

import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
BUILD_SCRIPT = ROOT / "build_windows_exe.py"


def _load_build_module():
    spec = importlib.util.spec_from_file_location("build_windows_exe", BUILD_SCRIPT)
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules["build_windows_exe"] = mod
    spec.loader.exec_module(mod)
    return mod


class TestBuildWindowsExeBundle(unittest.TestCase):
    def test_bundle_includes_changelog_md(self):
        mod = _load_build_module()
        self.assertIn("ЖУРНАЛ_ДОРАБОТОК.md", mod.BUNDLE_FILES)
        src = mod._bundle_src("ЖУРНАЛ_ДОРАБОТОК.md")
        self.assertTrue(src.is_file(), msg=f"Нет исходника {src}")

    def test_pyinstaller_includes_fonttools_for_fpdf(self):
        mod = _load_build_module()
        self.assertIn("fontTools", mod._PYI_COLLECT_SUBMODULES)
        self.assertIn("fontTools.varLib.iup", mod._PYI_EXTRA_HIDDEN)

    def test_pyinstaller_includes_overlapped_for_fpdf_asyncio(self):
        mod = _load_build_module()
        self.assertIn("_overlapped", mod._PYI_EXTRA_HIDDEN)

    def test_copy_bundle_asset_copies_file(self):
        mod = _load_build_module()
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "src.txt"
            dst = Path(tmp) / "dst.txt"
            src.write_text("hello", encoding="utf-8")
            self.assertTrue(mod._copy_bundle_asset(src, dst))
            self.assertEqual(dst.read_text(encoding="utf-8"), "hello")

    def test_ensure_update_config_creates_only_when_missing(self):
        mod = _load_build_module()
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp)
            self.assertTrue(mod._ensure_update_config(out))
            config_path = out / "update_config.json"
            self.assertTrue(config_path.is_file())
            self.assertFalse(mod._ensure_update_config(out))

    def test_local_update_share_dir(self):
        mod = _load_build_module()
        self.assertEqual(mod.LOCAL_UPDATE_SHARE_DIR, mod.ROOT / "UPDATE")

    def test_publish_local_update_share_writes_manifest(self):
        mod = _load_build_module()
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            exe = root / "ProtocolOOT.exe"
            exe.write_bytes(b"exe-payload")
            data_dir = root / "data"
            data_dir.mkdir()
            (data_dir / "FAQ.txt").write_text("faq", encoding="utf-8")
            update_dir = root / "UPDATE"
            original = mod.LOCAL_UPDATE_SHARE_DIR
            try:
                mod.LOCAL_UPDATE_SHARE_DIR = update_dir
                manifest_path = mod._publish_local_update_share(exe=exe, data_dir=data_dir)
            finally:
                mod.LOCAL_UPDATE_SHARE_DIR = original
            self.assertTrue(manifest_path.is_file())
            self.assertTrue((update_dir / "windows").exists())


if __name__ == "__main__":
    unittest.main()
