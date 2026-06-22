# -*- coding: utf-8 -*-
"""Инициализация SQLite (журнал, кэш Excel, настройки)."""

from __future__ import annotations

import sqlite3
from pathlib import Path

from commission_admin import (
    ensure_app_settings_table,
    ensure_commission_profiles_table,
    migrate_legacy_commission_profiles,
)
from excel_data_cache import ensure_excel_cache_tables
from protocol_docx import ensure_v_prof_cache_table
from protocol_journal import purge_duplicate_protocol_journal_rows
from protocol_paths import database_path


def init_protocols_db_file(db_path: Path) -> int:
    """Создаёт или обновляет схему SQLite; возвращает число удалённых дублей журнала."""
    db_path = Path(db_path)
    db_path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS protocols (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fio TEXT NOT NULL,
                topic TEXT,
                date TEXT,
                grade TEXT,
                content TEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        try:
            conn.execute("ALTER TABLE protocols ADD COLUMN export_meta_json TEXT")
        except sqlite3.OperationalError:
            pass
        try:
            conn.execute("ALTER TABLE protocols ADD COLUMN protocol_kind TEXT")
        except sqlite3.OperationalError:
            pass
        ensure_v_prof_cache_table(conn)
        ensure_excel_cache_tables(conn)
        ensure_app_settings_table(conn)
        ensure_commission_profiles_table(conn)
        migrate_legacy_commission_profiles(conn)
        conn.commit()
    return purge_duplicate_protocol_journal_rows(db_path)


def init_db() -> int:
    """Инициализация БД; возвращает число удалённых дублей в журнале протоколов."""
    return init_protocols_db_file(database_path())
