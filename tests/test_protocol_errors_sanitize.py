# -*- coding: utf-8 -*-
"""Обезличивание ПДн в журнале ошибок."""

from __future__ import annotations

import unittest

from _bootstrap import setup_main_project_paths

setup_main_project_paths()

from protocol_errors import sanitize_error_journal_text  # noqa: E402


class TestSanitizeErrorJournal(unittest.TestCase):
    def test_redacts_snils(self) -> None:
        raw = "Ошибка для 123-456-789 00"
        self.assertIn("[СНИЛС]", sanitize_error_journal_text(raw))
        self.assertNotIn("123-456-789", sanitize_error_journal_text(raw))

    def test_redacts_fio_label_line(self) -> None:
        raw = "Сотрудник: Иванов Иван Иванович\nНомер: 1"
        out = sanitize_error_journal_text(raw)
        self.assertIn("[обезличено]", out)
        self.assertNotIn("Иванов", out)
        self.assertIn("Номер: 1", out)


if __name__ == "__main__":
    unittest.main()
