# -*- coding: utf-8 -*-
"""Склонение фраз в родительный падеж (подстановки комиссии). Нужен pymorphy2 + pymorphy2-dicts-ru."""

from __future__ import annotations

import re
from typing import Any

_morph_analyzer: Any = None
_morph_failed: bool = False


def _get_morph():
    """Один MorphAnalyzer на процесс; при отсутствии пакета — None."""
    global _morph_analyzer, _morph_failed
    if _morph_failed:
        return None
    if _morph_analyzer is not None:
        return _morph_analyzer
    try:
        import pymorphy2

        _morph_analyzer = pymorphy2.MorphAnalyzer()
    except ImportError:
        _morph_failed = True
        return None
    return _morph_analyzer


def _strip_edges_punct(token: str) -> tuple[str, str, str]:
    """Скобки/кавычки слева, пунктуация справа — отдельно от слова для разбора."""
    lead = ""
    trail = ""
    t = token
    while t and t[0] in "(\"'«„":
        lead += t[0]
        t = t[1:]
    while t and t[-1] in ")\"'».…,:;!?":
        trail = t[-1] + trail
        t = t[:-1]
    return lead, t, trail


def _skip_word_for_inflect(mid: str) -> bool:
    if not mid:
        return True
    if re.fullmatch(r"[\d\.\-]+$", mid):
        return True
    # Инициалы вида И.О. или И. О.
    if re.fullmatch(r"([А-ЯЁA-Z]\.)+[А-ЯЁA-Z]?\.?", mid, re.IGNORECASE):
        return True
    if len(mid) <= 2 and "." in mid:
        return True
    return False


def _restore_case(original: str, inflected: str) -> str:
    if not inflected:
        return inflected
    if original.isupper():
        return inflected.upper()
    if len(original) >= 2 and original[0].isupper() and original[1:].islower():
        return inflected[0].upper() + inflected[1:] if len(inflected) > 1 else inflected.upper()
    if original[0].isupper():
        return inflected[0].upper() + inflected[1:]
    return inflected


def phrase_to_genitive_russian(phrase: str) -> str:
    """
    По «словам» (фрагменты между пробелами) — родительный падеж.
    Без pymorphy2 возвращает исходную строку.
    """
    phrase = (phrase or "").strip()
    if not phrase:
        return phrase
    morph = _get_morph()
    if morph is None:
        return phrase

    out: list[str] = []
    for part in re.split(r"(\s+)", phrase):
        if not part:
            continue
        if part.isspace():
            out.append(part)
            continue
        lead, mid, trail = _strip_edges_punct(part)
        if _skip_word_for_inflect(mid):
            out.append(part)
            continue
        parsed = morph.parse(mid)
        if not parsed:
            out.append(part)
            continue
        inf = parsed[0].inflect({"gent"})
        if not inf:
            out.append(part)
            continue
        w = _restore_case(mid, inf.word)
        out.append(f"{lead}{w}{trail}")
    return "".join(out)


def format_person_fio_profession_genitive(fio: str, profession: str) -> str:
    f = (fio or "").strip()
    p = (profession or "").strip()
    fg = phrase_to_genitive_russian(f) if f else ""
    pg = phrase_to_genitive_russian(p) if p else ""
    if fg and pg:
        return f"{fg}, {pg}"
    return fg or pg
