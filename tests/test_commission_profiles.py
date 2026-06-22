# -*- coding: utf-8 -*-
"""Профили комиссий по названию подразделения."""

from __future__ import annotations

import sqlite3
import tempfile
import unittest
from pathlib import Path

from _bootstrap import setup_main_project_paths

setup_main_project_paths()

from commission_admin import (
    COMMISSION_KIND_OT,
    delete_commission_profile,
    ensure_commission_profiles_table,
    list_commission_profile_names,
    load_commission_profile,
    save_commission_profile,
)
from employees_io import EmployeeRecord
from protocol_db import init_protocols_db_file


class TestCommissionProfiles(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.db_path = Path(self._tmp.name) / "protocols.db"
        init_protocols_db_file(self.db_path)

        import commission_admin as ca

        self._orig = ca.database_path
        ca.database_path = lambda: self.db_path  # type: ignore[assignment]
        with sqlite3.connect(self.db_path) as conn:
            ensure_commission_profiles_table(conn)
            conn.commit()

    def tearDown(self) -> None:
        import gc

        import commission_admin as ca

        ca.database_path = self._orig  # type: ignore[assignment]
        gc.collect()
        self._tmp.cleanup()

    def test_save_load_and_list_profiles(self) -> None:
        chair = EmployeeRecord(fio="Петров П.П.", profession="инженер")
        save_commission_profile(
            "НПС и ЦТТ",
            COMMISSION_KIND_OT,
            order_no="125",
            order_date="01.01.2026",
            chair=chair,
            members=[],
            venue_subdivision="НПС",
            order_approver="директор",
        )
        names = list_commission_profile_names(COMMISSION_KIND_OT)
        self.assertIn("НПС и ЦТТ", names)
        prof = load_commission_profile("НПС и ЦТТ", COMMISSION_KIND_OT)
        self.assertIsNotNone(prof)
        assert prof is not None
        self.assertEqual(prof["order_no"], "125")
        self.assertEqual(prof["chair"].fio, "Петров П.П.")

        save_commission_profile(
            "НПС и ЦТТ",
            COMMISSION_KIND_OT,
            order_no="126",
            order_date="02.02.2026",
            chair=chair,
            members=[],
            venue_subdivision="ЦТТ",
            order_approver="директор",
        )
        prof2 = load_commission_profile("НПС и ЦТТ", COMMISSION_KIND_OT)
        assert prof2 is not None
        self.assertEqual(prof2["order_no"], "126")
        self.assertEqual(len(list_commission_profile_names(COMMISSION_KIND_OT)), 1)

        delete_commission_profile("НПС и ЦТТ", COMMISSION_KIND_OT)
        self.assertEqual(list_commission_profile_names(COMMISSION_KIND_OT), [])


if __name__ == "__main__":
    unittest.main()
