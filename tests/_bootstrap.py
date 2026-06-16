# -*- coding: utf-8 -*-
"""Общая настройка sys.path для unit-тестов."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
_NEXT = ROOT / "ProtocolOHT_next"
_LINUX_PORT = ROOT / "linux_port"
_APP = _LINUX_PORT / "app"


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
