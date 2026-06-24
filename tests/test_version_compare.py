# -*- coding: utf-8
"""Сравнение версий приложения."""

from __future__ import annotations

import unittest

from _bootstrap import setup_main_project_paths

setup_main_project_paths()

from version_compare import is_newer_version, parse_version  # noqa: E402


class TestVersionCompare(unittest.TestCase):
    def test_parse_version_three_parts(self) -> None:
        self.assertEqual(parse_version("1.2.3"), (1, 2, 3))

    def test_parse_version_two_parts_as_patch_zero(self) -> None:
        self.assertEqual(parse_version("1.5"), (1, 5, 0))
        self.assertEqual(parse_version("1.5.0"), (1, 5, 0))

    def test_is_newer_version(self) -> None:
        self.assertTrue(is_newer_version("1.5.1", "1.5"))
        self.assertTrue(is_newer_version("1.6.0", "1.5.1"))
        self.assertFalse(is_newer_version("1.5", "1.5.0"))
        self.assertFalse(is_newer_version("1.5.0", "1.5.1"))


if __name__ == "__main__":
    unittest.main()
