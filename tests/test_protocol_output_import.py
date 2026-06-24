# -*- coding: utf-8
"""Импорт protocol_output без загрузки fpdf при старте."""

from __future__ import annotations

import importlib
import sys
import unittest

from _bootstrap import setup_main_project_paths

setup_main_project_paths()


class TestProtocolOutputImport(unittest.TestCase):
    def tearDown(self) -> None:
        for name in list(sys.modules):
            if name == "protocol_output" or name.startswith("fpdf"):
                del sys.modules[name]

    def test_import_does_not_load_fpdf(self) -> None:
        for name in list(sys.modules):
            if name == "protocol_output" or name.startswith("fpdf"):
                del sys.modules[name]
        import protocol_output  # noqa: F401

        self.assertNotIn("fpdf", sys.modules)
        importlib.reload(protocol_output)
        self.assertNotIn("fpdf", sys.modules)


if __name__ == "__main__":
    unittest.main()
