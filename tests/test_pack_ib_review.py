# -*- coding: utf-8
"""Пакет исходников для проверки ИБ."""

from __future__ import annotations

import importlib.util
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PACK = ROOT / "tools" / "pack_ib_review.py"


def _load_pack():
    spec = importlib.util.spec_from_file_location("pack_ib_review", PACK)
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules["pack_ib_review"] = mod
    spec.loader.exec_module(mod)
    return mod


class TestPackIbReview(unittest.TestCase):
    def test_pack_includes_bundle_integration_and_manifest(self) -> None:
        mod = _load_pack()
        with tempfile.TemporaryDirectory() as tmp:
            zpath = Path(tmp) / "ib_test.zip"
            mod.pack(out_zip=zpath)
            self.assertTrue(zpath.is_file())
            with zipfile.ZipFile(zpath) as zf:
                names = set(zf.namelist())
            self.assertIn("bundle_integration.py", names)
            self.assertIn("ИБ_MANIFEST.txt", names)
            self.assertIn("tests/test_pack_ib_review.py", names)
            self.assertIn("linux_port/prepare.py", names)
            self.assertNotIn("linux_port/app/main.py", names)


if __name__ == "__main__":
    unittest.main()
