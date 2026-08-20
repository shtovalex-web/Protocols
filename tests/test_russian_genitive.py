# -*- coding: utf-8 -*-
"""Родительный падеж для шапки комиссии в протоколе."""

from __future__ import annotations

import unittest

from _bootstrap import setup_main_project_paths

setup_main_project_paths()

from russian_genitive import (  # noqa: E402
    format_person_fio_profession_genitive,
    morph_backend_name,
    phrase_to_genitive_russian,
)


class TestRussianGenitive(unittest.TestCase):
    def test_morph_backend_available(self) -> None:
        name = morph_backend_name()
        self.assertIn(name, ("pymorphy3", "pymorphy2"), msg=f"нет морфологии: {name!r}")

    def test_surname_to_genitive(self) -> None:
        if morph_backend_name() == "none":
            self.skipTest("pymorphy не установлен")
        g = phrase_to_genitive_russian("Иванов")
        self.assertEqual(g, "Иванова")

    def test_commission_person_line_genitive(self) -> None:
        """Как в шапке после «членов комиссии»: ФИО и должность в род. п."""
        if morph_backend_name() == "none":
            self.skipTest("pymorphy не установлен")
        line = format_person_fio_profession_genitive(
            "Иванов Иван Иванович",
            "инженер",
        )
        self.assertIn("Иванова", line)
        self.assertIn("Ивана", line)
        self.assertIn("Ивановича", line)
        self.assertIn("инженера", line)
        self.assertTrue(line.startswith("Иванова"))


if __name__ == "__main__":
    unittest.main()
