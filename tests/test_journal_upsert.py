# -*- coding: utf-8 -*-
"""Журнал протоколов: перезапись записи вместо дублирования."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from _bootstrap import (
    close_tracked_sqlite_connections,
    enable_sqlite_test_tracking,
    setup_main_project_paths,
)

setup_main_project_paths()

from protocol_db import init_protocols_db_file
from employees_io import EmployeeRecord
from protocol_journal import (
    PROTOCOL_JOURNAL_KIND_OT,
    build_protocol_export_meta_json,
    dedupe_journal_records_for_export,
    delete_protocol_journal_rows_by_ids,
    get_all_protocols,
    get_protocols_journal_display,
    purge_duplicate_protocol_journal_rows,
    save_protocol,
)


class TestJournalUpsert(unittest.TestCase):
    def setUp(self) -> None:
        enable_sqlite_test_tracking()
        self._tmp = tempfile.TemporaryDirectory()
        self.db_path = Path(self._tmp.name) / "protocols.db"
        init_protocols_db_file(self.db_path)

        import protocol_journal as pj

        self._orig = pj.database_path
        pj.database_path = lambda: self.db_path  # type: ignore[assignment]

    def tearDown(self) -> None:
        import protocol_journal as pj

        pj.database_path = self._orig  # type: ignore[assignment]
        close_tracked_sqlite_connections()
        self._tmp.cleanup()

    def test_save_protocol_overwrites_same_number_date_kind(self) -> None:
        meta1 = build_protocol_export_meta_json(
            ["B"],
            ["Базовая"],
            protocol_no_formatted="5-06-2026",
            persons_raw=[EmployeeRecord(fio="Иванов И.И.", profession="Слесарь")],
            face_sheet_profession="Слесарь",
        )
        id1 = save_protocol(
            "Иванов И.И.",
            "Базовая",
            "09.06.2026",
            "удовлетворительно",
            "",
            export_meta_json=meta1,
            protocol_kind=PROTOCOL_JOURNAL_KIND_OT,
        )
        meta2 = build_protocol_export_meta_json(
            ["B", "PP"],
            ["Базовая", "Первая помощь"],
            protocol_no_formatted="5-06-2026",
            persons_raw=[EmployeeRecord(fio="Иванов И.И.", profession="Слесарь")],
            face_sheet_profession="Слесарь",
        )
        id2 = save_protocol(
            "Иванов И.И.",
            "Базовая; Первая помощь",
            "09.06.2026",
            "неудовлетворительно",
            "",
            export_meta_json=meta2,
            protocol_kind=PROTOCOL_JOURNAL_KIND_OT,
        )
        self.assertEqual(id1, id2)
        rows = get_all_protocols(PROTOCOL_JOURNAL_KIND_OT)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["grade"], "неудовлетворительно")
        self.assertIn("Первая помощь", rows[0]["topic"] or "")

    def test_save_protocol_new_row_same_number_different_profession(self) -> None:
        meta1 = build_protocol_export_meta_json(
            ["B"],
            ["Базовая"],
            protocol_no_formatted="5-06-2026",
            persons_raw=[EmployeeRecord(fio="Иванов И.И.", profession="Слесарь")],
            face_sheet_profession="Слесарь",
        )
        id1 = save_protocol(
            "Иванов И.И.",
            "Базовая — слесарь",
            "09.06.2026",
            "удовлетворительно",
            "",
            export_meta_json=meta1,
            protocol_kind=PROTOCOL_JOURNAL_KIND_OT,
        )
        meta2 = build_protocol_export_meta_json(
            ["B"],
            ["Базовая"],
            protocol_no_formatted="5-06-2026",
            persons_raw=[EmployeeRecord(fio="Иванов И.И.", profession="Электрик")],
            face_sheet_profession="Электрик",
        )
        id2 = save_protocol(
            "Иванов И.И.",
            "Базовая — электрик",
            "09.06.2026",
            "удовлетворительно",
            "",
            export_meta_json=meta2,
            protocol_kind=PROTOCOL_JOURNAL_KIND_OT,
        )
        self.assertNotEqual(id1, id2)
        rows = get_all_protocols(PROTOCOL_JOURNAL_KIND_OT)
        self.assertEqual(len(rows), 2)
        topics = sorted((r.get("topic") or "") for r in rows)
        self.assertEqual(topics, ["Базовая — слесарь", "Базовая — электрик"])

    def test_save_protocol_new_row_same_number_profession_different_fio(self) -> None:
        """Один № и должность, разные ФИО — две записи (не перезаписывать чужую)."""
        meta1 = build_protocol_export_meta_json(
            ["B"],
            ["Базовая"],
            protocol_no_formatted="8-06-2026",
            persons_raw=[EmployeeRecord(fio="Иванов И.И.", profession="Слесарь")],
            face_sheet_profession="Слесарь",
        )
        id1 = save_protocol(
            "Иванов И.И.",
            "Базовая",
            "09.06.2026",
            "удовлетворительно",
            "",
            export_meta_json=meta1,
            protocol_kind=PROTOCOL_JOURNAL_KIND_OT,
        )
        meta2 = build_protocol_export_meta_json(
            ["B"],
            ["Базовая"],
            protocol_no_formatted="8-06-2026",
            persons_raw=[EmployeeRecord(fio="Петров П.П.", profession="Слесарь")],
            face_sheet_profession="Слесарь",
        )
        id2 = save_protocol(
            "Петров П.П.",
            "Базовая",
            "09.06.2026",
            "удовлетворительно",
            "",
            export_meta_json=meta2,
            protocol_kind=PROTOCOL_JOURNAL_KIND_OT,
        )
        self.assertNotEqual(id1, id2)
        rows = get_all_protocols(PROTOCOL_JOURNAL_KIND_OT)
        self.assertEqual(len(rows), 2)
        fios = sorted((r.get("fio") or "") for r in rows)
        self.assertEqual(fios, ["Иванов И.И.", "Петров П.П."])

    def test_dedupe_keeps_distinct_profession_same_protocol_no(self) -> None:
        meta1 = build_protocol_export_meta_json(
            ["B"],
            ["Базовая"],
            protocol_no_formatted="7-06-2026",
            face_sheet_profession="Слесарь",
        )
        meta2 = build_protocol_export_meta_json(
            ["B"],
            ["Базовая"],
            protocol_no_formatted="7-06-2026",
            face_sheet_profession="Электрик",
        )
        save_protocol(
            "Иванов И.И.",
            "A",
            "09.06.2026",
            "удовлетворительно",
            "",
            export_meta_json=meta1,
            protocol_kind=PROTOCOL_JOURNAL_KIND_OT,
        )
        save_protocol(
            "Иванов И.И.",
            "B",
            "09.06.2026",
            "удовлетворительно",
            "",
            export_meta_json=meta2,
            protocol_kind=PROTOCOL_JOURNAL_KIND_OT,
        )
        rows = dedupe_journal_records_for_export(get_all_protocols(PROTOCOL_JOURNAL_KIND_OT))
        self.assertEqual(len(rows), 2)

    def test_delete_protocol_journal_rows_by_ids(self) -> None:
        id1 = save_protocol(
            "Иванов И.И.",
            "A",
            "09.06.2026",
            "удовлетворительно",
            "",
            protocol_kind=PROTOCOL_JOURNAL_KIND_OT,
        )
        id2 = save_protocol(
            "Петров П.П.",
            "B",
            "09.06.2026",
            "удовлетворительно",
            "",
            protocol_kind=PROTOCOL_JOURNAL_KIND_OT,
        )
        removed = delete_protocol_journal_rows_by_ids([id1])
        self.assertEqual(removed, 1)
        rows = get_all_protocols(PROTOCOL_JOURNAL_KIND_OT)
        self.assertEqual(len(rows), 1)
        self.assertEqual(int(rows[0]["id"]), id2)

    def test_dedupe_keeps_distinct_fio_without_protocol_no(self) -> None:
        save_protocol("Иванов И.И.", "A", "09.06.2026", "удовлетворительно", "")
        save_protocol("Петров П.П.", "B", "09.06.2026", "удовлетворительно", "")
        rows = dedupe_journal_records_for_export(get_all_protocols(PROTOCOL_JOURNAL_KIND_OT))
        self.assertEqual(len(rows), 2)

    def test_get_protocols_journal_display_hides_duplicates(self) -> None:
        meta = build_protocol_export_meta_json(
            ["B"],
            ["Базовая"],
            protocol_no_formatted="2-06-2026",
            face_sheet_profession="Слесарь",
        )
        import sqlite3

        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                INSERT INTO protocols (fio, topic, date, grade, content, export_meta_json, protocol_kind)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                ("Иванов И.И.", "старая", "09.06.2026", "удовлетворительно", "", meta, PROTOCOL_JOURNAL_KIND_OT),
            )
            conn.commit()
        save_protocol(
            "Иванов И.И.",
            "новая",
            "09.06.2026",
            "неудовлетворительно",
            "",
            export_meta_json=meta,
            protocol_kind=PROTOCOL_JOURNAL_KIND_OT,
        )
        shown = get_protocols_journal_display(PROTOCOL_JOURNAL_KIND_OT)
        self.assertEqual(len(shown), 1)
        self.assertEqual(shown[0]["topic"], "новая")

    def test_purge_removes_old_duplicate_rows(self) -> None:
        meta = build_protocol_export_meta_json(
            ["B"],
            ["Базовая"],
            protocol_no_formatted="1-06-2026",
            face_sheet_profession="Слесарь",
        )
        import sqlite3

        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                INSERT INTO protocols (fio, topic, date, grade, content, export_meta_json, protocol_kind)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    "Иванов И.И.",
                    "старая",
                    "09.06.2026",
                    "удовлетворительно",
                    "",
                    meta,
                    PROTOCOL_JOURNAL_KIND_OT,
                ),
            )
            conn.execute(
                """
                INSERT INTO protocols (fio, topic, date, grade, content, export_meta_json, protocol_kind)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    "Иванов И.И.",
                    "новая",
                    "09.06.2026",
                    "неудовлетворительно",
                    "",
                    meta,
                    PROTOCOL_JOURNAL_KIND_OT,
                ),
            )
            conn.commit()
        removed = purge_duplicate_protocol_journal_rows(self.db_path)
        self.assertEqual(removed, 1)
        rows = get_all_protocols(PROTOCOL_JOURNAL_KIND_OT)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["topic"], "новая")


if __name__ == "__main__":
    unittest.main()
