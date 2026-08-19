# -*- coding: utf-8
"""Выгрузка Минтруд: слияние строк без дублирования."""

from __future__ import annotations

import unittest

from _bootstrap import setup_main_project_paths

setup_main_project_paths()

from mintrud_export import dedupe_export_data_rows, merge_export_data_rows


class TestMintrudExportMerge(unittest.TestCase):
    def test_merge_overwrites_same_person_program_protocol(self) -> None:
        old = [
            {
                "protocol_no": "5-06-2026",
                "doc_date": "2026-06-09",
                "last_name": "Иванов",
                "first_name": "Иван",
                "patronymic": "Иванович",
                "program_name": "Базовая",
                "test_passed": 1,
            }
        ]
        new = [
            {
                "protocol_no": "5-06-2026",
                "doc_date": "2026-06-09",
                "last_name": "Иванов",
                "first_name": "Иван",
                "patronymic": "Иванович",
                "program_name": "Базовая",
                "test_passed": 0,
            }
        ]
        merged = merge_export_data_rows(old, new)
        self.assertEqual(len(merged), 1)
        self.assertEqual(merged[0]["test_passed"], 0)

    def test_dedupe_export_rows(self) -> None:
        rows = [
            {
                "protocol_no": "1",
                "doc_date": "2026-01-01",
                "last_name": "A",
                "first_name": "B",
                "patronymic": "",
                "program_name": "P",
            },
            {
                "protocol_no": "1",
                "doc_date": "2026-01-01",
                "last_name": "A",
                "first_name": "B",
                "patronymic": "",
                "program_name": "P",
            },
        ]
        self.assertEqual(len(dedupe_export_data_rows(rows)), 1)

    def test_merge_keeps_same_program_different_position(self) -> None:
        old = [
            {
                "protocol_no": "5-06-2026",
                "doc_date": "2026-06-09",
                "last_name": "Иванов",
                "first_name": "Иван",
                "patronymic": "Иванович",
                "program_name": "Базовая",
                "position": "Слесарь",
                "test_passed": 1,
            }
        ]
        new = [
            {
                "protocol_no": "5-06-2026",
                "doc_date": "2026-06-09",
                "last_name": "Иванов",
                "first_name": "Иван",
                "patronymic": "Иванович",
                "program_name": "Базовая",
                "position": "Электрик",
                "test_passed": 1,
            }
        ]
        merged = merge_export_data_rows(old, new)
        self.assertEqual(len(merged), 2)
        positions = sorted(r.get("position") or "" for r in merged)
        self.assertEqual(positions, ["Слесарь", "Электрик"])


if __name__ == "__main__":
    unittest.main()
