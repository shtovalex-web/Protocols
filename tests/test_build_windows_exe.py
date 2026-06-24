# -*- coding: utf-8 -*-
"""Комплект Windows-сборки (build_windows_exe.py)."""

from __future__ import annotations

import importlib.util
import sys
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


if __name__ == "__main__":
    unittest.main()
