# -*- coding: utf-8 -*-
"""Склонение фраз в родительный падеж (подстановки комиссии).

Предпочтительно pymorphy3 (+ pymorphy3-dicts-ru) — работает на Python 3.11+.
Запасной вариант: pymorphy2 (+ pymorphy2-dicts-ru); на новых Python часто
нужен setuptools (pkg_resources). Без морфологии строки возвращаются как есть.
"""

from __future__ import annotations

import re
from typing import Any

_morph_analyzer: Any = None
_morph_backend: str | None = None
_morph_failed: bool = False
_morph_missing_warned: bool = False
_morph_init_error: str = ""

# После этих предлогов остаток фразы не склоняем (типично для должностей:
# «инженер по охране труда» → «инженера по охране труда»).
_PREPOSITIONS = frozenset(
    {
        "по",
        "в",
        "во",
        "на",
        "о",
        "об",
        "обо",
        "от",
        "до",
        "для",
        "при",
        "с",
        "со",
        "к",
        "ко",
        "у",
        "из",
        "без",
        "над",
        "под",
        "перед",
        "между",
        "через",
        "про",
        "за",
    }
)


def morph_backend_name() -> str:
    """Имя активного бэкенда: pymorphy3 | pymorphy2 | none."""
    _get_morph()
    if _morph_failed or _morph_backend is None:
        return "none"
    return _morph_backend


def maybe_warn_missing_morphology(parent: Any = None) -> bool:
    """
    Один раз за процесс предупредить, если нет pymorphy3/pymorphy2.
    Возвращает True, если морфология доступна.
    """
    global _morph_missing_warned
    if morph_backend_name() != "none":
        return True
    if _morph_missing_warned:
        return False
    _morph_missing_warned = True
    try:
        import sys
        from tkinter import messagebox

        if getattr(sys, "frozen", False):
            hint = (
                "В этой сборке нет модуля склонения.\n"
                "Обратитесь к администратору за обновлённой версией программы."
            )
        else:
            # Подсказка для того же интерпретатора, которым запущена программа.
            py = sys.executable or "py -3"
            hint = (
                f"Установка для текущего Python:\n"
                f'"{py}" -m pip install pymorphy3 pymorphy3-dicts-ru'
            )
            if _morph_init_error:
                hint += f"\n\nТехническая причина: {_morph_init_error}"

        messagebox.showwarning(
            "Родительный падеж",
            "Не удалось подключить библиотеки склонения (pymorphy3 / pymorphy2).\n"
            "ФИО и должности комиссии в протоколе останутся в исходном падеже.\n\n"
            + hint,
            parent=parent,
        )
    except Exception:
        pass
    return False


def _get_morph():
    """Один MorphAnalyzer на процесс; при отсутствии пакетов — None."""
    global _morph_analyzer, _morph_backend, _morph_failed, _morph_init_error
    if _morph_failed:
        return None
    if _morph_analyzer is not None:
        return _morph_analyzer

    errors: list[str] = []

    # pymorphy3 — основной для Python 3.11+ (в т.ч. 3.12–3.14).
    try:
        import pymorphy3

        _morph_analyzer = pymorphy3.MorphAnalyzer()
        _morph_backend = "pymorphy3"
        _morph_init_error = ""
        return _morph_analyzer
    except Exception as e:
        errors.append(f"pymorphy3: {type(e).__name__}: {e}")

    try:
        import pymorphy2

        _morph_analyzer = pymorphy2.MorphAnalyzer()
        _morph_backend = "pymorphy2"
        _morph_init_error = ""
        return _morph_analyzer
    except Exception as e:
        errors.append(f"pymorphy2: {type(e).__name__}: {e}")
        _morph_failed = True
        _morph_backend = None
        _morph_init_error = "; ".join(errors)
        return None


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


def _tag_has(tag: Any, key: str) -> bool:
    try:
        return key in tag
    except Exception:
        return False


def _detect_fio_gender(morph: Any, phrase: str) -> str | None:
    """Определить пол по имени/отчеству в ФИО: 'femn' | 'masc' | None."""
    for part in re.split(r"\s+", (phrase or "").strip()):
        if not part:
            continue
        _lead, mid, _trail = _strip_edges_punct(part)
        if _skip_word_for_inflect(mid) or "-" in mid:
            continue
        for p in morph.parse(mid)[:4]:
            if not (_tag_has(p.tag, "Name") or _tag_has(p.tag, "Patr")):
                continue
            if _tag_has(p.tag, "femn"):
                return "femn"
            if _tag_has(p.tag, "masc"):
                return "masc"
    return None


def _pick_parse(morph: Any, mid: str, *, gender: str | None) -> Any | None:
    parsed = morph.parse(mid)
    if not parsed:
        return None
    if gender == "femn":
        for p in parsed:
            if _tag_has(p.tag, "femn") and _tag_has(p.tag, "Surn"):
                return p
        for p in parsed:
            if _tag_has(p.tag, "femn"):
                return p
    elif gender == "masc":
        for p in parsed:
            if _tag_has(p.tag, "masc") and _tag_has(p.tag, "Surn"):
                return p
        for p in parsed:
            if _tag_has(p.tag, "masc") and (
                _tag_has(p.tag, "Name") or _tag_has(p.tag, "Patr") or _tag_has(p.tag, "Surn")
            ):
                return p
    return parsed[0]


def _inflect_token_to_genitive(morph: Any, mid: str, *, gender: str | None) -> str:
    if "-" in mid:
        parts = mid.split("-")
        out_parts: list[str] = []
        for part in parts:
            if not part:
                out_parts.append(part)
                continue
            out_parts.append(_inflect_simple_word(morph, part, gender=gender))
        return "-".join(out_parts)
    return _inflect_simple_word(morph, mid, gender=gender)


def _inflect_simple_word(morph: Any, mid: str, *, gender: str | None) -> str:
    parsed = _pick_parse(morph, mid, gender=gender)
    if parsed is None:
        return mid
    inf = parsed.inflect({"gent"})
    if not inf:
        return mid
    return _restore_case(mid, inf.word)


def phrase_to_genitive_russian(phrase: str, *, gender: str | None = None) -> str:
    """
    По «словам» (фрагменты между пробелами) — родительный падеж.
    Без морфологического пакета возвращает исходную строку.

    gender: подсказка пола для фамилий ('femn'/'masc'); если None — определяется
    по имени/отчеству в этой же фразе (если есть).
    После предлога (по, в, на, …) остальные слова не склоняются.
    """
    phrase = (phrase or "").strip()
    if not phrase:
        return phrase
    morph = _get_morph()
    if morph is None:
        return phrase

    hint = gender if gender in ("femn", "masc") else _detect_fio_gender(morph, phrase)

    out: list[str] = []
    stop_inflect = False
    for part in re.split(r"(\s+)", phrase):
        if not part:
            continue
        if part.isspace():
            out.append(part)
            continue
        lead, mid, trail = _strip_edges_punct(part)
        if stop_inflect or _skip_word_for_inflect(mid):
            out.append(part)
            continue
        if mid.casefold() in _PREPOSITIONS:
            out.append(part)
            stop_inflect = True
            continue
        w = _inflect_token_to_genitive(morph, mid, gender=hint)
        out.append(f"{lead}{w}{trail}")
    return "".join(out)


def format_person_fio_profession_genitive(fio: str, profession: str) -> str:
    """Строка для шапки протокола: «ФИО, должность» в родительном падеже."""
    f = (fio or "").strip()
    p = (profession or "").strip()
    morph = _get_morph()
    gender = _detect_fio_gender(morph, f) if morph is not None and f else None
    fg = phrase_to_genitive_russian(f, gender=gender) if f else ""
    # Должность склоняем без подсказки пола ФИО (иначе «слесарь» может поехать).
    pg = phrase_to_genitive_russian(p, gender=None) if p else ""
    if fg and pg:
        return f"{fg}, {pg}"
    return fg or pg
