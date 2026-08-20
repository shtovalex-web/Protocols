# -*- coding: utf-8 -*-
"""Родительный падеж для шапки комиссии в протоколе."""

from __future__ import annotations

import unittest

from _bootstrap import setup_main_project_paths

setup_main_project_paths()

from commission_admin import apply_commission_insertions_to_line  # noqa: E402
from russian_genitive import (  # noqa: E402
    format_person_fio_profession_genitive,
    morph_backend_name,
    phrase_to_genitive_russian,
)


class TestRussianGenitive(unittest.TestCase):
    def setUp(self) -> None:
        if morph_backend_name() == "none":
            self.skipTest("pymorphy не установлен")

    def test_morph_backend_available(self) -> None:
        name = morph_backend_name()
        self.assertIn(name, ("pymorphy3", "pymorphy2"), msg=f"нет морфологии: {name!r}")

    def test_surname_to_genitive(self) -> None:
        g = phrase_to_genitive_russian("Иванов")
        self.assertEqual(g, "Иванова")

    def test_commission_person_line_genitive(self) -> None:
        """Как в шапке после «членов комиссии»: ФИО и должность в род. п."""
        line = format_person_fio_profession_genitive(
            "Иванов Иван Иванович",
            "инженер",
        )
        self.assertIn("Иванова", line)
        self.assertIn("Ивана", line)
        self.assertIn("Ивановича", line)
        self.assertIn("инженера", line)
        self.assertTrue(line.startswith("Иванова"))

    def test_female_fio_genitive(self) -> None:
        """Женская фамилия не должна оставаться как мужской род. п. «Петрова»."""
        line = format_person_fio_profession_genitive(
            "Петрова Анна Сергеевна",
            "инженер",
        )
        self.assertIn("Петровой", line)
        self.assertIn("Анны", line)
        self.assertIn("Сергеевны", line)
        self.assertNotEqual(line.split(",")[0].strip(), "Петрова Анна Сергеевна")

    def test_profession_keeps_prepositional_tail(self) -> None:
        """После предлога «по» хвост должности не склоняем в род. п."""
        g = phrase_to_genitive_russian("инженер по охране труда")
        self.assertEqual(g, "инженера по охране труда")

    def test_hyphenated_surname_keeps_capitals(self) -> None:
        g = phrase_to_genitive_russian("Кузнецов-Смирнов")
        self.assertEqual(g, "Кузнецова-Смирнова")

    def test_initials_unchanged(self) -> None:
        g = phrase_to_genitive_russian("Сидоров А.А.")
        self.assertEqual(g, "Сидорова А.А.")


class TestCommissionHeaderInsertions(unittest.TestCase):
    def setUp(self) -> None:
        if morph_backend_name() == "none":
            self.skipTest("pymorphy не установлен")

    def test_separate_lines_chair_and_members(self) -> None:
        chair = format_person_fio_profession_genitive(
            "Иванов Иван Иванович", "инженер"
        )
        members = format_person_fio_profession_genitive(
            "Петрова Анна Сергеевна", "слесарь"
        )
        out_ch = apply_commission_insertions_to_line(
            "председателя",
            date_words="",
            order_no="",
            chair_gen=chair,
            members_gen=members,
        )
        out_m = apply_commission_insertions_to_line(
            "членов комиссии",
            date_words="",
            order_no="",
            chair_gen=chair,
            members_gen=members,
        )
        self.assertIn(chair, out_ch)
        self.assertIn(members, out_m)
        self.assertIn("Петровой", out_m)

    def test_same_paragraph_keeps_chair_and_members(self) -> None:
        """Один абзац с «председателя» и «членов» — не терять председателя при оформлении."""
        from docx import Document

        from protocol_docx import _apply_commission_paragraph_replacement

        chair = "Иванова Ивана Ивановича, инженера"
        members = "Петровой Анны Сергеевны, слесаря"
        orig = (
            "комиссии в составе: председателя _________________ "
            "и членов _________________"
        )
        new = apply_commission_insertions_to_line(
            orig,
            date_words="",
            order_no="",
            chair_gen=chair,
            members_gen=members,
        )
        self.assertIn(chair, new)
        self.assertIn(members, new)

        doc = Document()
        para = doc.add_paragraph(orig)
        _apply_commission_paragraph_replacement(
            para,
            orig,
            new,
            members_gen=members,
            chair_gen=chair,
        )
        text = para.text
        self.assertIn("Иванова", text)
        self.assertIn("Петровой", text)


if __name__ == "__main__":
    unittest.main()
