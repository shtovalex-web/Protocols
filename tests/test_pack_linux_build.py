# -*- coding: utf-8 -*-
"""Тест формирования автономного комплекта Linux-сборки."""

from __future__ import annotations

import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

from pack_linux_build import pack  # noqa: E402


class TestPackLinuxBuild(unittest.TestCase):
    def test_pack_creates_build_kit(self):
        if not (ROOT / "linux_port" / "release" / "build_release_linux.py").is_file():
            self.skipTest("нет linux_port/release/")

        with tempfile.TemporaryDirectory(prefix="protocoloot_kit_") as tmpdir:
            out = Path(tmpdir) / "kit"
            try:
                pack(out, skip_prepare=True)
            except SystemExit as e:
                if not (ROOT / "linux_port" / "app" / "main.py").is_file():
                    self.skipTest("нет linux_port/app — выполните prepare.py")
                raise

            self.assertTrue((out / "app" / "main.py").is_file())
            self.assertTrue((out / "release" / "build_release_linux.py").is_file())
            self.assertTrue((out / "check_env.sh").is_file())
            self.assertTrue((out / "build.sh").is_file())
            self.assertTrue((out / "requirements-build.txt").is_file())
            self.assertTrue((out / "VERSION.txt").is_file())
            text = (out / "requirements-build.txt").read_text(encoding="utf-8")
            self.assertIn("-r requirements.txt", text)
            for sh in out.glob("*.sh"):
                self.assertNotIn(b"\r", sh.read_bytes(), msg=f"CRLF в {sh.name}")
            self.assertTrue((out / "fix_crlf.py").is_file())
            self.assertTrue((out / "lib" / "sh_common.sh").is_file())


if __name__ == "__main__":
    unittest.main()
