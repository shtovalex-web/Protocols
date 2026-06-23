# -*- coding: utf-8 -*-
"""Общая настройка sys.path для unit-тестов."""

from __future__ import annotations

import gc
import sqlite3
import sys
from collections.abc import Callable
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
_NEXT = ROOT / "ProtocolOHT_next"
_LINUX_PORT = ROOT / "linux_port"
_APP = _LINUX_PORT / "app"

_orig_sqlite_connect: Callable[..., sqlite3.Connection] = sqlite3.connect
_tracked_sqlite_connections: list[sqlite3.Connection] = []
_sqlite_tracking_enabled = False


def setup_main_project_paths() -> None:
    for _p in (str(_NEXT), str(ROOT)):
        while _p in sys.path:
            sys.path.remove(_p)
    sys.path.insert(0, str(_NEXT))
    sys.path.insert(1, str(ROOT))


def setup_linux_port_app_paths() -> None:
    if not (_APP / "main.py").is_file():
        raise RuntimeError("Нет linux_port/app — выполните: python linux_port/prepare.py")
    for _p in (str(_APP), str(_APP / "ProtocolOHT_next")):
        while _p in sys.path:
            sys.path.remove(_p)
    sys.path.insert(0, str(_APP / "ProtocolOHT_next"))
    sys.path.insert(1, str(_APP))


def enable_sqlite_test_tracking() -> None:
    """Учитывать соединения sqlite3.connect (контекстный менеджер их не закрывает)."""
    global _sqlite_tracking_enabled
    if _sqlite_tracking_enabled:
        return
    _sqlite_tracking_enabled = True

    def _tracking_connect(*args: Any, **kwargs: Any) -> sqlite3.Connection:
        conn = _orig_sqlite_connect(*args, **kwargs)
        _tracked_sqlite_connections.append(conn)
        return conn

    sqlite3.connect = _tracking_connect  # type: ignore[method-assign]


def close_tracked_sqlite_connections() -> None:
    """Закрыть учтённые соединения перед удалением тестовой БД (Windows)."""
    global _sqlite_tracking_enabled
    while _tracked_sqlite_connections:
        conn = _tracked_sqlite_connections.pop()
        try:
            conn.close()
        except sqlite3.Error:
            pass
    if _sqlite_tracking_enabled:
        sqlite3.connect = _orig_sqlite_connect  # type: ignore[method-assign]
        _sqlite_tracking_enabled = False
    gc.collect()
