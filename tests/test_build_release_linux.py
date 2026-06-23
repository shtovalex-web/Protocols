# -*- coding: utf-8 -*-
"""Тест комплекта Linux-сборки (assemble_release)."""

from __future__ import annotations

import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
BUILD_SCRIPT = ROOT / "linux_port" / "release" / "build_release_linux.py"


def _load_build_module():
    spec = importlib.util.spec_from_file_location("build_release_linux", BUILD_SCRIPT)
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules["build_release_linux"] = mod
    spec.loader.exec_module(mod)
    return mod


class TestAssembleReleaseLinux(unittest.TestCase):
    def test_assemble_release_creates_package(self):
        mod = _load_build_module()
        instruction = BUILD_SCRIPT.parent / mod.INSTRUCTION_NAME
        self.assertTrue(instruction.is_file(), msg=f"Нет {instruction}")

        with patch.object(mod, "_data_subdir_name", return_value="data"):
            with tempfile.TemporaryDirectory() as tmpdir:
                tmp = Path(tmpdir)
                fake_binary = tmp / "src.bin"
                fake_binary.write_bytes(b"\x7fELF-fake")

                result = mod.assemble_release(tmp / "out", binary_source=fake_binary)

                self.assertEqual(result["binary"].name, mod.BINARY_NAME)
                self.assertEqual(result["binary"].read_bytes(), b"\x7fELF-fake")
                self.assertTrue(result["instruction"].is_file())
                self.assertTrue((result["data"] / "README.txt").is_file())


if __name__ == "__main__":
    unittest.main()
