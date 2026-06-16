# -*- coding: utf-8 -*-
"""Ключи программ протокола (листы B, PP, SIZ, V — журнал, Минтруд, таблица результатов)."""

from __future__ import annotations

from enum import Enum
from typing import Literal, Sequence


class ProgramKey(str, Enum):
    """Ключ программы в export_meta_json и в логике сборки протокола."""

    B = "B"
    PP = "PP"
    SIZ = "SIZ"
    V = "V"


ProgramKeyLiteral = Literal["B", "PP", "SIZ", "V"]


def parse_program_key(raw: str | None) -> ProgramKey | None:
    """Распознаёт ключ; регистр для латиницы не важен (b → B)."""
    s = (raw or "").strip()
    if not s:
        return None
    u = s.upper()
    for member in ProgramKey:
        if member.value == u:
            return member
    return None


def program_key_to_str(key: ProgramKey | str) -> str:
    return key.value if isinstance(key, ProgramKey) else str(key)


def program_keys_as_str_list(keys: Sequence[ProgramKey | str] | None) -> list[str]:
    if not keys:
        return []
    return [program_key_to_str(k) for k in keys]


def program_keys_contains(keys: Sequence[ProgramKey | str] | None, want: ProgramKey) -> bool:
    if not keys:
        return False
    for k in keys:
        if isinstance(k, ProgramKey):
            if k == want:
                return True
            continue
        if parse_program_key(str(k)) == want:
            return True
    return False
