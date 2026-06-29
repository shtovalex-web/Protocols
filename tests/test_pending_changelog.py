# -*- coding: utf-8
"""Отложенный changelog после обновления."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from _bootstrap import setup_main_project_paths

setup_main_project_paths()

from pending_changelog import pop_pending_changelog, write_pending_changelog  # noqa: E402


class TestPendingChangelog(unittest.TestCase):
    def test_write_and_pop_pending_changelog(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            data_dir = Path(tmp)
            write_pending_changelog(
                data_dir,
                version="1.5.5",
                changes=["Первое", "Второе"],
            )
            pending = pop_pending_changelog(data_dir)
            self.assertEqual(pending, ("1.5.5", ["Первое", "Второе"]))
            self.assertIsNone(pop_pending_changelog(data_dir))

    def test_pop_missing_returns_none(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            self.assertIsNone(pop_pending_changelog(Path(tmp)))


if __name__ == "__main__":
    unittest.main()
