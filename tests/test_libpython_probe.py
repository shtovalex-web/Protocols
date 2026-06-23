# -*- coding: utf-8 -*-
"""Проверка libpython_probe.py."""

from __future__ import annotations

import subprocess
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PROBE = ROOT / "linux_port" / "lib" / "libpython_probe.py"


class TestLibpythonProbe(unittest.TestCase):
    def test_probe_runs_with_current_python(self) -> None:
        if not PROBE.is_file():
            self.skipTest("нет libpython_probe.py")
        r = subprocess.run(
            [sys.executable, str(PROBE)],
            capture_output=True,
            text=True,
        )
        # На CI/Windows без dev-пакетов может быть 1 — главное, скрипт не падает.
        self.assertIn(r.returncode, (0, 1))

    def test_verbose_prints_path_when_ok(self) -> None:
        if not PROBE.is_file():
            self.skipTest("нет libpython_probe.py")
        r = subprocess.run(
            [sys.executable, str(PROBE), "--verbose"],
            capture_output=True,
            text=True,
        )
        if r.returncode == 0:
            self.assertTrue(r.stdout.strip())


if __name__ == "__main__":
    unittest.main()
