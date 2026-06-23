# -*- coding: utf-8 -*-
"""Должности для поиска в V_PROF при ручной подстановке в поле «Должность»."""

from __future__ import annotations

import unittest

from _bootstrap import setup_main_project_paths

setup_main_project_paths()

from employees_io import EmployeeRecord
from protocol_docx import collect_professions_for_v_prof_lookup


class TestCollectProfessionsForVProfLookup(unittest.TestCase):
    def test_face_sheet_override_skips_excel_primary(self) -> None:
        rec = EmployeeRecord(
            fio="Иванов И.И.",
            profession="Оператор технологических установок",
            profession2="",
        )
        manual = "Слесарь по ремонту технологических установок"
        result = collect_professions_for_v_prof_lookup(
            face_sheet_profession=manual,
            persons_merged=[rec],
            persons_row_source=[rec],
        )
        self.assertEqual(result, [manual])

    def test_face_sheet_override_keeps_profession2(self) -> None:
        rec = EmployeeRecord(
            fio="Иванов И.И.",
            profession="Оператор",
            profession2="Электромонтер",
        )
        manual = "Слесарь по ремонту технологических установок"
        result = collect_professions_for_v_prof_lookup(
            face_sheet_profession=manual,
            persons_merged=[rec],
            persons_row_source=[rec],
        )
        self.assertEqual(result, [manual, "Электромонтер"])

    def test_without_face_sheet_uses_excel_primary(self) -> None:
        rec = EmployeeRecord(
            fio="Иванов И.И.",
            profession="Оператор",
            profession2="Электромонтер",
        )
        result = collect_professions_for_v_prof_lookup(
            persons_merged=[rec],
            persons_row_source=[rec],
        )
        self.assertEqual(result, ["Оператор", "Электромонтер"])


if __name__ == "__main__":
    unittest.main()
