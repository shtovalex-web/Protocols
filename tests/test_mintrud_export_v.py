# -*- coding: utf-8 -*-
"""Выгрузка Минтруд: программа «В» — все должности и программы из протокола."""

from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from _bootstrap import setup_main_project_paths

setup_main_project_paths()

from employees_io import EmployeeRecord
from mintrud_export import build_export_rows, mintrud_v_program_entries_for_employee
from protocol_journal import build_protocol_export_meta_json


class TestMintrudExportV(unittest.TestCase):
    def _record_with_meta(self, meta_json: str) -> dict:
        return {
            "id": 1,
            "date": "15.01.2026",
            "grade": "удовлетворительно",
            "export_meta_json": meta_json,
            "fio": "Иванов Иван Иванович",
            "content": "",
        }

    def test_build_export_rows_v_creates_row_per_protocol_program_with_position(self) -> None:
        meta = build_protocol_export_meta_json(
            ["V"],
            ["Программа (В)"],
            protocol_no_formatted="1-01-26",
            persons_raw=[
                EmployeeRecord(
                    fio="Иванов Иван Иванович",
                    profession="Слесарь",
                    profession2="Электрик",
                )
            ],
            persons_row_source=[
                EmployeeRecord(fio="Иванов Иван Иванович", profession="Слесарь"),
                EmployeeRecord(fio="Иванов Иван Иванович", profession="Электрик"),
            ],
        )
        entries = [
            ("2.1 Общие требования", "Слесарь"),
            ("3.1 Электробезопасность", "Электрик"),
        ]
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            catalog = Path(tmp.name)
        try:
            with mock.patch(
                "mintrud_export.get_cached_v_registry_rows",
                return_value=[],
            ):
                with mock.patch(
                    "mintrud_export.mintrud_v_program_entries_for_employee",
                    return_value=entries,
                ):
                    rows = build_export_rows(
                        [self._record_with_meta(meta)],
                        inn_employer="123",
                        employer_name="ООО Тест",
                        programs_excel_path=catalog,
                    )
        finally:
            catalog.unlink(missing_ok=True)
        self.assertEqual(len(rows), 2)
        self.assertEqual(
            {r["program_name"] for r in rows},
            {"2.1 Общие требования", "3.1 Электробезопасность"},
        )
        self.assertEqual({r["position"] for r in rows}, {"Слесарь", "Электрик"})

    def test_build_export_rows_b_keeps_lookup_position_not_protocol_only(self) -> None:
        meta = build_protocol_export_meta_json(
            ["B"],
            ["Программа Б"],
            protocol_no_formatted="1-01-26",
            persons_raw=[
                EmployeeRecord(fio="Иванов Иван Иванович", profession="Слесарь"),
            ],
        )
        employees = [
            EmployeeRecord(
                fio="Иванов Иван Иванович",
                profession="Должность из Excel",
                snils="123-456-789 00",
            )
        ]
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            catalog = Path(tmp.name)
        try:
            with mock.patch(
                "mintrud_export.get_cached_b_program_title",
                return_value="Программа общая",
            ):
                with mock.patch(
                    "mintrud_export.get_cached_v_registry_rows",
                    return_value=[("1", "Программа общая", "", None)],
                ):
                    rows = build_export_rows(
                        [self._record_with_meta(meta)],
                        employees=employees,
                        programs_excel_path=catalog,
                    )
        finally:
            catalog.unlink(missing_ok=True)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["position"], "Должность из Excel")

    def test_mintrud_v_entries_uses_protocol_context(self) -> None:
        meta = json.loads(
            build_protocol_export_meta_json(
                ["V"],
                ["Программа (В)"],
                protocol_no_formatted="1-01-26",
                persons_raw=[
                    EmployeeRecord(fio="Петров П.П.", profession="Слесарь"),
                ],
                persons_row_source=[
                    EmployeeRecord(fio="Петров П.П.", profession="Слесарь"),
                    EmployeeRecord(fio="Петров П.П.", profession="Электрик"),
                ],
                face_sheet_profession="Слесарь",
                v_prof_enabled_by_fio={
                    "петров п.п.": frozenset({"слесарь", "электрик"}),
                },
                v_prof_main_by_fio={"петров п.п.": "Слесарь"},
            )
        )
        emp = EmployeeRecord(fio="Петров П.П.")
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            catalog = Path(tmp.name)
        try:
            with mock.patch(
                "protocol_docx.v_program_parts_with_professions_for_employee",
                return_value=[
                    ("Фрагмент 1", "Слесарь"),
                    ("Фрагмент 2", "Электрик"),
                ],
            ) as mocked:
                result = mintrud_v_program_entries_for_employee(
                    catalog,
                    emp,
                    meta,
                    employees=None,
                    lookup={},
                )
            self.assertEqual(len(result), 2)
            mocked.assert_called_once()
            kwargs = mocked.call_args.kwargs
            self.assertEqual(len(kwargs["persons_row_source"]), 2)
            self.assertEqual(kwargs["face_sheet_profession"], "Слесарь")
        finally:
            catalog.unlink(missing_ok=True)


if __name__ == "__main__":
    unittest.main()
