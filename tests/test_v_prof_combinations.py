# -*- coding: utf-8 -*-
"""Совмещения должностей для программ «В»."""

from __future__ import annotations

import tkinter as tk
import unittest

from _bootstrap import setup_main_project_paths

setup_main_project_paths()

from employees_io import EmployeeRecord
from v_prof_combinations import (
    VProfCombinationsDialog,
    needs_combinations_dialog,
    professions_by_fio,
)


class TestVProfCombinationsLogic(unittest.TestCase):
    def test_professions_by_fio_merges_same_fio_rows(self) -> None:
        recs = [
            EmployeeRecord(fio="Иванов И.И.", profession="Оператор", profession2=""),
            EmployeeRecord(fio="Иванов И.И.", profession="Электромонтер", profession2=""),
        ]
        groups = professions_by_fio(recs)
        self.assertEqual(len(groups), 1)
        self.assertEqual(groups[0][1], "Иванов И.И.")
        self.assertEqual(len(groups[0][2]), 2)

    def test_needs_combinations_when_profession_and_profession2(self) -> None:
        rec = EmployeeRecord(
            fio="Петров П.П.",
            profession="Слесарь",
            profession2="Электромонтер",
        )
        self.assertTrue(needs_combinations_dialog([rec]))

    def test_no_combinations_for_single_profession(self) -> None:
        rec = EmployeeRecord(fio="Сидоров С.С.", profession="Оператор", profession2="")
        self.assertFalse(needs_combinations_dialog([rec]))


class TestVProfCombinationsDialog(unittest.TestCase):
    def test_dialog_builds_checkboxes_for_professions(self) -> None:
        rec = EmployeeRecord(
            fio="Иванов И.И.",
            profession="Оператор",
            profession2="Электромонтер",
        )
        groups = professions_by_fio([rec])
        root = tk.Tk()
        root.withdraw()
        try:
            dlg = VProfCombinationsDialog(root, groups)
            dlg.update_idletasks()
            self.assertEqual(len(dlg._main_vars), 1)
            self.assertEqual(len(dlg._check_vars), 2)
            children = [w for w in dlg.winfo_children()]
            self.assertTrue(children)
        finally:
            try:
                dlg.destroy()
            except NameError:
                pass
            root.destroy()


if __name__ == "__main__":
    unittest.main()
