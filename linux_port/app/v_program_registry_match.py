# -*- coding: utf-8 -*-
"""Сопоставление текста с колонкой B листа V (фрагмент V_PROF или заголовок Б/PP/СИЗ): C/B — наименование, A — ID."""

from __future__ import annotations

import re
from collections.abc import Sequence


def norm_profession_key(s: str) -> str:
    return re.sub(r"\s+", " ", s.strip().lower().replace("ё", "е"))


def fg_line_comparison_key(s: str) -> str:
    t = s.replace("\xa0", " ").strip()
    t = re.sub(r"^[-–—]\s*", "", t)
    return norm_profession_key(t)


def match_v_registry_fragment(
    fragment: str,
    rows: Sequence[tuple[str, str, str] | tuple[str, str, str, float | None]],
) -> tuple[str, str, float | None] | None:
    """
    Фрагмент V_PROF ↔ столбец B листа V.
    Возвращает (наименование для реестра из C или B, ID из A, часы из столбца D листа V) или None.
    Строка кэша может быть (A,B,C) или (A,B,C,hours).
    """
    raw = (fragment or "").strip()
    f = fg_line_comparison_key(fragment)
    if not f:
        return None
    if raw.isdigit() or re.fullmatch(r"\d+(?:[.,]\d+)?", raw):
        try:
            pid = str(int(float(raw.replace(",", "."))))
        except ValueError:
            pid = raw
        for row in rows:
            gos = str(row[0]).strip()
            try:
                gos_n = str(int(float(gos)))
            except (TypeError, ValueError):
                gos_n = gos
            if gos_n != pid:
                continue
            title = (str(row[2]).strip() if len(row) > 2 and row[2] else "") or (
                str(row[1]).strip() if len(row) > 1 and row[1] else ""
            )
            if not title:
                continue
            h_val: float | None = None
            if len(row) > 3 and row[3] is not None:
                try:
                    h_val = float(row[3])
                except (TypeError, ValueError):
                    h_val = None
            return (title, gos, h_val)
    if len(f) < 2:
        return None
    best: tuple[int, int, str, str, float | None] | None = None
    for row in rows:
        if len(row) >= 4:
            gos_id, b_raw, c_raw, h_raw = row[0], row[1], row[2], row[3]
        else:
            gos_id, b_raw, c_raw = row[0], row[1], row[2]
            h_raw = None
        nb = fg_line_comparison_key(b_raw)
        if not nb:
            continue
        score = 0
        if f == nb:
            score = 3
        elif len(nb) >= 4 and (f in nb or nb in f):
            score = 2
        elif f in nb or nb in f:
            score = 1
        if score == 0:
            continue
        title = (c_raw or b_raw).strip()
        if not title:
            continue
        h_val: float | None
        if h_raw is None:
            h_val = None
        else:
            try:
                h_val = float(h_raw)
            except (TypeError, ValueError):
                h_val = None
        cand = (score, len(nb), title, gos_id, h_val)
        if best is None or cand[0] > best[0] or (
            cand[0] == best[0] and cand[1] > best[1]
        ):
            best = cand
    if best is None:
        return None
    return (best[2], best[3], best[4])


def _v_row_by_program_id(
    program_id: int | str,
    rows: Sequence[tuple[str, str, str] | tuple[str, str, str, float | None]],
) -> tuple[str, str, str] | tuple[str, str, str, float | None] | None:
    pid = str(program_id).strip()
    try:
        pid_norm = str(int(float(pid)))
    except (TypeError, ValueError):
        pid_norm = pid
    for row in rows:
        gos = str(row[0]).strip()
        try:
            gos_norm = str(int(float(gos)))
        except (TypeError, ValueError):
            gos_norm = gos
        if gos_norm == pid_norm:
            return row
    return None


def header_title_for_v_program_id(
    program_id: int | str,
    rows: Sequence[tuple[str, str, str] | tuple[str, str, str, float | None]],
) -> str | None:
    """Шапка протокола и абзацы после проверки: только столбец C листа V (по ID в A)."""
    row = _v_row_by_program_id(program_id, rows)
    if row is None:
        m = match_v_registry_fragment(str(program_id), rows)
        if m:
            return (m[0] or "").strip() or None
        return None
    c_raw = row[2] if len(row) > 2 else ""
    return (str(c_raw).strip() if c_raw else "") or None


def table_text_for_v_program_id(
    program_id: int | str,
    rows: Sequence[tuple[str, str, str] | tuple[str, str, str, float | None]],
) -> str | None:
    """Таблица протокола и Минтруд «В»: только столбец B листа V (по ID в A)."""
    row = _v_row_by_program_id(program_id, rows)
    if row is None:
        return None
    b_raw = row[1] if len(row) > 1 else ""
    return (str(b_raw).strip() if b_raw else "") or None


def table_text_for_v_prof_fragment(
    fragment: str,
    rows: Sequence[tuple[str, str, str] | tuple[str, str, str, float | None]],
) -> str:
    """
    Текст для таблицы/Минтруд из столбца B листа V при совпадении фрагмента с B;
    иначе исходный фрагмент (legacy: ячейка V_PROF).
    """
    raw = (fragment or "").strip()
    if not raw or not rows:
        return raw
    f = fg_line_comparison_key(raw)
    best_b = ""
    best_score = 0
    for row in rows:
        b_raw = row[1] if len(row) > 1 else ""
        nb = fg_line_comparison_key(b_raw)
        if not nb:
            continue
        score = 0
        if f == nb:
            score = 3
        elif len(nb) >= 4 and (f in nb or nb in f):
            score = 2
        elif f in nb or nb in f:
            score = 1
        if score > best_score:
            best_score = score
            best_b = str(b_raw).strip()
    return best_b if best_score > 0 else raw


def title_for_v_program_id(
    program_id: int | str,
    rows: Sequence[tuple[str, str, str] | tuple[str, str, str, float | None]],
) -> str | None:
    """Совместимость: то же, что header_title_for_v_program_id (столбец C)."""
    return header_title_for_v_program_id(program_id, rows)
