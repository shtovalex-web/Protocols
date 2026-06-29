# -*- coding: utf-8 -*-
"""Комплект Windows-сборки (build_windows_exe.py)."""

from __future__ import annotations

import importlib.util
import json
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

    def test_ensure_update_config_points_to_deploy_share_for_onefile(self):
        mod = _load_build_module()
        with tempfile.TemporaryDirectory() as tmp:
            out = mod.DEFAULT_OUT_DIR
            if out.is_dir():
                config_path = out / "update_config.json"
                if config_path.is_file():
                    payload = json.loads(config_path.read_text(encoding="utf-8"))
                    self.assertIn("Обновление", payload["manifest_path"])
                    return
            out = Path(tmp) / "onefile"
            out.mkdir()
            original = mod.DEFAULT_OUT_DIR
            try:
                mod.DEFAULT_OUT_DIR = out
                self.assertTrue(mod._ensure_update_config(out))
                payload = json.loads((out / "update_config.json").read_text(encoding="utf-8"))
                self.assertIn("Обновление", payload["manifest_path"])
            finally:
                mod.DEFAULT_OUT_DIR = original

    def test_ensure_update_config_creates_only_when_missing(self):
        mod = _load_build_module()
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "custom"
            out.mkdir()
            self.assertTrue(mod._ensure_update_config(out))
            config_path = out / "update_config.json"
            self.assertTrue(config_path.is_file())
            self.assertFalse(mod._ensure_update_config(out))

    def test_deploy_update_share_dir(self):
        mod = _load_build_module()
        self.assertEqual(mod.DEPLOY_UPDATE_SHARE_DIR, Path(r"D:\Обновление"))

    def test_write_update_info_creates_json(self):
        mod = _load_build_module()
        with tempfile.TemporaryDirectory() as tmp:
            data_dir = Path(tmp)
            path = mod._write_update_info(data_dir)
            self.assertTrue(path.is_file())
            payload = json.loads(path.read_text(encoding="utf-8"))
            self.assertTrue(payload["version"])

    def test_publish_update_share_writes_manifest(self):
        mod = _load_build_module()
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            exe = root / "ProtocolOOT.exe"
            exe.write_bytes(b"exe-payload")
            data_dir = root / "data"
            data_dir.mkdir()
            mod._write_update_info(data_dir)
            (data_dir / "FAQ.txt").write_text("faq", encoding="utf-8")
            share_dir = root / "deploy_share"
            manifest_path = mod._publish_update_share(
                exe=exe,
                data_dir=data_dir,
                share_root=share_dir,
            )
            self.assertTrue(manifest_path.is_file())
            self.assertTrue((share_dir / "windows").exists())

    def test_try_publish_deploy_returns_none_on_failure(self):
        mod = _load_build_module()
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            data_dir = root / "data"
            data_dir.mkdir()
            missing_exe = root / "missing.exe"
            original = mod.DEPLOY_UPDATE_SHARE_DIR
            try:
                mod.DEPLOY_UPDATE_SHARE_DIR = root / "deploy"
                result = mod._try_publish_deploy_update_share(
                    exe=missing_exe,
                    data_dir=data_dir,
                )
            finally:
                mod.DEPLOY_UPDATE_SHARE_DIR = original
            self.assertIsNone(result)


if __name__ == "__main__":
    unittest.main()
