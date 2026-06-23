# -*- coding: utf-8 -*-
"""Smoke-тест сборки протокола в DOCX по шаблону из bundle/."""

from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path

from _bootstrap import setup_main_project_paths

setup_main_project_paths()

from employees_io import EmployeeRecord
from protocol_docx import (
    build_filled_protocol_document,
    format_protocol_number_for_template,
    is_word_protocol_template,
    protocol_template_path,
)


class TestProtocolDocxSmoke(unittest.TestCase):
    def test_build_filled_protocol_document_writes_valid_docx(self) -> None:
        template = protocol_template_path()
        self.assertTrue(
            is_word_protocol_template(template),
            f"Нет шаблона протокола: {template}",
        )

        person = EmployeeRecord(
            fio="Иванов Иван Иванович",
            profession="Слесарь",
            subdivision="Цех 1",
        )
        protocol_no = "1"
        date_str = "15.06.2026"
        program_title = "Базовая программа обучения"

        doc, table_excess = build_filled_protocol_document(
            template,
            protocol_no=protocol_no,
            date_str=date_str,
            theme=program_title,
            table_persons=[person],
            program_keys=["B"],
            program_titles=[program_title],
            grade="удовлетворительно",
        )
        self.assertGreaterEqual(table_excess, 0)
        self.assertGreater(len(doc.paragraphs), 0)

        expected_no = format_protocol_number_for_template(protocol_no, date_str)

        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "smoke_protocol.docx"
            doc.save(str(out))
            self.assertTrue(out.is_file())
            self.assertGreater(out.stat().st_size, 5000)

            with zipfile.ZipFile(out) as zf:
                names = zf.namelist()
                self.assertIn("word/document.xml", names)
                xml = zf.read("word/document.xml").decode("utf-8")
                if expected_no:
                    self.assertIn(expected_no, xml)
                self.assertIn("Иванов", xml)


if __name__ == "__main__":
    unittest.main()
