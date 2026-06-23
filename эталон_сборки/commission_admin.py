# -*- coding: utf-8 -*-
"""Приказ о комиссии и состав комиссии: SQLite (app_settings) и блок интерфейса."""

from __future__ import annotations

import json
import re
import sqlite3
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Callable

from russian_genitive import format_person_fio_profession_genitive

import tkinter as tk
from tkinter import messagebox
from tkinter import ttk

from ui_theme import (
    FIELD_COMBO_STYLE,
    FIELD_DATE_STYLE,
    FIELD_STYLE,
    configure_editable_text,
    configure_listbox,
    pad,
)

from app_paths import application_user_dir
from clipboard_ui import bind_editable_clipboard, register_clipboard_window
from employees_io import (
    EmployeeExcelError,
    EmployeeRecord,
    employee_unique_key,
    format_fio_iof,
    listbox_label_for_employee,
    load_commission_from_excel,
    sort_employees_by_fio_alphabet,
)

DATABASE_FILENAME = "protocols.db"

COMMISSION_KIND_OT = "ot"
COMMISSION_KIND_TECH = "tech"

SETTING_COMMISSION_ORDER_NO = "commission_order_no"
SETTING_COMMISSION_ORDER_DATE = "commission_order_date"
SETTING_COMMISSION_CHAIR_JSON = "commission_chair_json"
SETTING_COMMISSION_MEMBERS_JSON = "commission_members_json"

SETTING_TECH_COMMISSION_ORDER_NO = "tech_commission_order_no"
SETTING_TECH_COMMISSION_ORDER_DATE = "tech_commission_order_date"
SETTING_TECH_COMMISSION_CHAIR_JSON = "tech_commission_chair_json"
SETTING_TECH_COMMISSION_MEMBERS_JSON = "tech_commission_members_json"
SETTING_COMMISSION_VENUE_SUBDIVISION = "commission_venue_subdivision"
SETTING_COMMISSION_ORDER_APPROVER = "commission_order_approver"
SETTING_TECH_COMMISSION_VENUE_SUBDIVISION = "tech_commission_venue_subdivision"
SETTING_TECH_COMMISSION_ORDER_APPROVER = "tech_commission_order_approver"
SETTING_COMMISSION_ACTIVE_PROFILE_OT = "commission_active_profile_name_ot"
SETTING_COMMISSION_ACTIVE_PROFILE_TECH = "commission_active_profile_name_tech"
# Путь к .docx шаблону тех. протокола (пусто — default_protocol_tehnicheskiy.docx из папки программы).
SETTING_TECH_PROTOCOL_TEMPLATE_DOCX = "tech_protocol_template_docx"
# «1» — при запуске включать защиту стандартных шаблонов в папке программы (см. docx_template_protection).
SETTING_PROTECT_BUNDLE_TEMPLATES = "protect_bundle_protocol_templates"

# Плейсхолдеры в .docx (не удалять при правке шаблона).
COMMISSION_VENUE_PLACEHOLDER = "{{ПОДРАЗДЕЛЕНИЕ_ПРОВЕРКИ}}"
COMMISSION_ORDER_APPROVER_PLACEHOLDER = "{{УТВЕРДИЛ_ПРИКАЗ}}"


def load_tech_protocol_template_docx_path() -> str:
    return _app_setting_get(SETTING_TECH_PROTOCOL_TEMPLATE_DOCX, "")


def save_tech_protocol_template_docx_path(path: str) -> None:
    _app_setting_set(SETTING_TECH_PROTOCOL_TEMPLATE_DOCX, (path or "").strip())


def load_protect_bundle_templates_enabled() -> bool:
    return _app_setting_get(SETTING_PROTECT_BUNDLE_TEMPLATES, "1").strip() != "0"


def save_protect_bundle_templates_enabled(enabled: bool) -> None:
    _app_setting_set(SETTING_PROTECT_BUNDLE_TEMPLATES, "1" if enabled else "0")


def _commission_setting_keys(kind: str) -> tuple[str, str, str, str]:
    if kind == COMMISSION_KIND_TECH:
        return (
            SETTING_TECH_COMMISSION_ORDER_NO,
            SETTING_TECH_COMMISSION_ORDER_DATE,
            SETTING_TECH_COMMISSION_CHAIR_JSON,
            SETTING_TECH_COMMISSION_MEMBERS_JSON,
        )
    return (
        SETTING_COMMISSION_ORDER_NO,
        SETTING_COMMISSION_ORDER_DATE,
        SETTING_COMMISSION_CHAIR_JSON,
        SETTING_COMMISSION_MEMBERS_JSON,
    )


def _commission_context_setting_keys(kind: str) -> tuple[str, str]:
    if kind == COMMISSION_KIND_TECH:
        return (
            SETTING_TECH_COMMISSION_VENUE_SUBDIVISION,
            SETTING_TECH_COMMISSION_ORDER_APPROVER,
        )
    return (SETTING_COMMISSION_VENUE_SUBDIVISION, SETTING_COMMISSION_ORDER_APPROVER)

# Подписи в конце .docx (плейсхолдеры {{ПРЕДСЕДАТЕЛЬ}}, {{ЧЛЕНЫ_КОМИССИИ}})
_COMMISSION_SIG_UNDERSCORE_BEFORE = "_____"
_COMMISSION_SIG_UNDERSCORE_AFTER = "___________________"
_COMMISSION_SIG_CAPTION_INDENT = " " * 57
_COMMISSION_SIG_CAPTION_FALLBACK = "И.О. Фамилия, подпись"


def _commission_signature_caption_line(profession: str) -> str:
    p = (profession or "").strip()
    if p:
        return f"{_COMMISSION_SIG_CAPTION_INDENT}{p}, подпись"
    return f"{_COMMISSION_SIG_CAPTION_INDENT}{_COMMISSION_SIG_CAPTION_FALLBACK}"


def _format_chair_signature_block(fio: str, profession: str) -> str:
    iof = format_fio_iof(fio)
    if not iof:
        return ""
    line1 = (
        f"Председатель комиссии\t{_COMMISSION_SIG_UNDERSCORE_BEFORE}"
        f"{iof}{_COMMISSION_SIG_UNDERSCORE_AFTER}"
    )
    return f"{line1}\n{_commission_signature_caption_line(profession)}"


def _format_member_signature_block(fio: str, profession: str) -> str:
    iof = format_fio_iof(fio)
    if not iof:
        return ""
    line1 = (
        f"Члены комиссии:\t{_COMMISSION_SIG_UNDERSCORE_BEFORE}"
        f"{iof}{_COMMISSION_SIG_UNDERSCORE_AFTER}"
    )
    return f"{line1}\n{_commission_signature_caption_line(profession)}"


def database_path() -> Path:
    return application_user_dir() / DATABASE_FILENAME


def ensure_app_settings_table(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS app_settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )
        """
    )


def _app_setting_get(key: str, default: str = "") -> str:
    with sqlite3.connect(database_path()) as conn:
        ensure_app_settings_table(conn)
        row = conn.execute(
            "SELECT value FROM app_settings WHERE key = ?",
            (key,),
        ).fetchone()
        return str(row[0]) if row and row[0] is not None else default


def _app_setting_set(key: str, value: str) -> None:
    with sqlite3.connect(database_path()) as conn:
        ensure_app_settings_table(conn)
        conn.execute(
            """
            INSERT INTO app_settings (key, value) VALUES (?, ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value
            """,
            (key, value),
        )
        conn.commit()


def _employee_from_settings_dict(d: Any) -> EmployeeRecord:
    if not isinstance(d, dict):
        return EmployeeRecord(fio="")
    return EmployeeRecord(
        fio=str(d.get("fio", "")).strip(),
        profession=str(d.get("profession", "")).strip(),
        subdivision=str(d.get("subdivision", "")).strip(),
        profession2=str(d.get("profession2", "")).strip(),
        snils=str(d.get("snils", "")).strip(),
    )


def load_commission_state_from_db(
    kind: str = COMMISSION_KIND_OT,
) -> tuple[str, str, EmployeeRecord | None, list[EmployeeRecord]]:
    k0, k1, k2, k3 = _commission_setting_keys(kind)
    order_no = _app_setting_get(k0, "")
    order_date = _app_setting_get(k1, "")
    chair: EmployeeRecord | None = None
    raw_ch = _app_setting_get(k2, "")
    if raw_ch.strip():
        try:
            c = _employee_from_settings_dict(json.loads(raw_ch))
            if c.fio.strip():
                chair = c
        except (json.JSONDecodeError, TypeError, ValueError):
            chair = None
    members: list[EmployeeRecord] = []
    raw_m = _app_setting_get(k3, "")
    if raw_m.strip():
        try:
            arr = json.loads(raw_m)
            if isinstance(arr, list):
                for item in arr:
                    em = _employee_from_settings_dict(item)
                    if em.fio.strip():
                        members.append(em)
        except (json.JSONDecodeError, TypeError, ValueError):
            pass
    return order_no, order_date, chair, members


def load_commission_protocol_context_from_db(
    kind: str = COMMISSION_KIND_OT,
) -> tuple[str, str]:
    """Подразделение (место проверки) и фрагмент «кем утверждён приказ» для шаблона протокола."""
    k_venue, k_approver = _commission_context_setting_keys(kind)
    return (
        _app_setting_get(k_venue, "").strip(),
        _app_setting_get(k_approver, "").strip(),
    )


def save_commission_protocol_context_to_db(
    venue_subdivision: str,
    order_approver: str,
    *,
    kind: str = COMMISSION_KIND_OT,
) -> None:
    k_venue, k_approver = _commission_context_setting_keys(kind)
    _app_setting_set(k_venue, (venue_subdivision or "").strip())
    _app_setting_set(k_approver, (order_approver or "").strip())


def save_commission_state_to_db(
    order_no: str,
    order_date: str,
    chair: EmployeeRecord | None,
    members: list[EmployeeRecord],
    *,
    kind: str = COMMISSION_KIND_OT,
    venue_subdivision: str = "",
    order_approver: str = "",
) -> None:
    k0, k1, k2, k3 = _commission_setting_keys(kind)
    save_commission_protocol_context_to_db(
        venue_subdivision, order_approver, kind=kind
    )
    _app_setting_set(k0, (order_no or "").strip())
    _app_setting_set(k1, (order_date or "").strip())
    if chair is not None and (chair.fio or "").strip():
        _app_setting_set(
            k2,
            json.dumps(asdict(chair), ensure_ascii=False),
        )
    else:
        _app_setting_set(k2, "")
    _app_setting_set(
        k3,
        json.dumps(
            [asdict(m) for m in members if (m.fio or "").strip()],
            ensure_ascii=False,
        ),
    )


def ensure_commission_profiles_table(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS commission_profiles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            kind TEXT NOT NULL,
            order_no TEXT NOT NULL DEFAULT '',
            order_date TEXT NOT NULL DEFAULT '',
            venue_subdivision TEXT NOT NULL DEFAULT '',
            order_approver TEXT NOT NULL DEFAULT '',
            chair_json TEXT NOT NULL DEFAULT '',
            members_json TEXT NOT NULL DEFAULT '',
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(kind, name)
        )
        """
    )


def _active_commission_profile_setting_key(kind: str) -> str:
    if kind == COMMISSION_KIND_TECH:
        return SETTING_COMMISSION_ACTIVE_PROFILE_TECH
    return SETTING_COMMISSION_ACTIVE_PROFILE_OT


def set_active_commission_profile_name(name: str, *, kind: str = COMMISSION_KIND_OT) -> None:
    _app_setting_set(_active_commission_profile_setting_key(kind), (name or "").strip())


def get_active_commission_profile_name(kind: str = COMMISSION_KIND_OT) -> str:
    return _app_setting_get(_active_commission_profile_setting_key(kind), "").strip()


def list_commission_profile_names(kind: str = COMMISSION_KIND_OT) -> list[str]:
    k = (kind or COMMISSION_KIND_OT).strip() or COMMISSION_KIND_OT
    with sqlite3.connect(database_path()) as conn:
        ensure_commission_profiles_table(conn)
        cur = conn.execute(
            """
            SELECT name FROM commission_profiles
            WHERE kind = ?
            ORDER BY name COLLATE NOCASE
            """,
            (k,),
        )
        return [str(row[0]) for row in cur.fetchall() if row and row[0]]


def save_commission_profile(
    name: str,
    kind: str,
    *,
    order_no: str,
    order_date: str,
    chair: EmployeeRecord | None,
    members: list[EmployeeRecord],
    venue_subdivision: str = "",
    order_approver: str = "",
) -> None:
    profile_name = (name or "").strip()
    if not profile_name:
        raise ValueError("Укажите название комиссии (подразделение).")
    k = (kind or COMMISSION_KIND_OT).strip() or COMMISSION_KIND_OT
    chair_json = ""
    if chair is not None and (chair.fio or "").strip():
        chair_json = json.dumps(asdict(chair), ensure_ascii=False)
    members_json = json.dumps(
        [asdict(m) for m in members if (m.fio or "").strip()],
        ensure_ascii=False,
    )
    with sqlite3.connect(database_path()) as conn:
        ensure_commission_profiles_table(conn)
        conn.execute(
            """
            INSERT INTO commission_profiles (
                name, kind, order_no, order_date, venue_subdivision, order_approver,
                chair_json, members_json, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(kind, name) DO UPDATE SET
                order_no = excluded.order_no,
                order_date = excluded.order_date,
                venue_subdivision = excluded.venue_subdivision,
                order_approver = excluded.order_approver,
                chair_json = excluded.chair_json,
                members_json = excluded.members_json,
                updated_at = CURRENT_TIMESTAMP
            """,
            (
                profile_name,
                k,
                (order_no or "").strip(),
                (order_date or "").strip(),
                (venue_subdivision or "").strip(),
                (order_approver or "").strip(),
                chair_json,
                members_json,
            ),
        )
        conn.commit()
    set_active_commission_profile_name(profile_name, kind=k)


def load_commission_profile(
    name: str,
    kind: str = COMMISSION_KIND_OT,
) -> dict[str, Any] | None:
    profile_name = (name or "").strip()
    if not profile_name:
        return None
    k = (kind or COMMISSION_KIND_OT).strip() or COMMISSION_KIND_OT
    with sqlite3.connect(database_path()) as conn:
        ensure_commission_profiles_table(conn)
        row = conn.execute(
            """
            SELECT order_no, order_date, venue_subdivision, order_approver,
                   chair_json, members_json
            FROM commission_profiles
            WHERE kind = ? AND name = ?
            """,
            (k, profile_name),
        ).fetchone()
    if not row:
        return None
    chair: EmployeeRecord | None = None
    if (row[4] or "").strip():
        try:
            c = _employee_from_settings_dict(json.loads(row[4]))
            if c.fio.strip():
                chair = c
        except (json.JSONDecodeError, TypeError, ValueError):
            chair = None
    members: list[EmployeeRecord] = []
    if (row[5] or "").strip():
        try:
            arr = json.loads(row[5])
            if isinstance(arr, list):
                for item in arr:
                    em = _employee_from_settings_dict(item)
                    if em.fio.strip():
                        members.append(em)
        except (json.JSONDecodeError, TypeError, ValueError):
            pass
    return {
        "name": profile_name,
        "order_no": str(row[0] or ""),
        "order_date": str(row[1] or ""),
        "venue_subdivision": str(row[2] or ""),
        "order_approver": str(row[3] or ""),
        "chair": chair,
        "members": members,
    }


def delete_commission_profile(name: str, kind: str = COMMISSION_KIND_OT) -> None:
    profile_name = (name or "").strip()
    if not profile_name:
        return
    k = (kind or COMMISSION_KIND_OT).strip() or COMMISSION_KIND_OT
    with sqlite3.connect(database_path()) as conn:
        ensure_commission_profiles_table(conn)
        conn.execute(
            "DELETE FROM commission_profiles WHERE kind = ? AND name = ?",
            (k, profile_name),
        )
        conn.commit()
    if get_active_commission_profile_name(k) == profile_name:
        set_active_commission_profile_name("", kind=k)


def migrate_legacy_commission_profiles(conn: sqlite3.Connection) -> None:
    """Перенос единственной комиссии из app_settings в именованный профиль."""
    ensure_commission_profiles_table(conn)
    for kind in (COMMISSION_KIND_OT, COMMISSION_KIND_TECH):
        cnt = conn.execute(
            "SELECT COUNT(*) FROM commission_profiles WHERE kind = ?",
            (kind,),
        ).fetchone()
        if cnt and int(cnt[0]) > 0:
            continue
        order_no, order_date, chair, members = load_commission_state_from_db(kind)
        venue, approver = load_commission_protocol_context_from_db(kind)
        if not any(
            [
                (order_no or "").strip(),
                (order_date or "").strip(),
                chair,
                members,
                (venue or "").strip(),
                (approver or "").strip(),
            ]
        ):
            continue
        default_name = (venue.splitlines()[0].strip() if venue.strip() else "") or "Основная"
        try:
            save_commission_profile(
                default_name,
                kind,
                order_no=order_no,
                order_date=order_date,
                chair=chair,
                members=members,
                venue_subdivision=venue,
                order_approver=approver,
            )
        except ValueError:
            pass
        if not get_active_commission_profile_name(kind):
            set_active_commission_profile_name(default_name, kind=kind)


SETTING_MINTRUD_INN_EMPLOYER = "mintrud_inn_employer"
SETTING_MINTRUD_EMPLOYER_NAME = "mintrud_employer_name"
SETTING_MINTRUD_INN_ORG2 = "mintrud_inn_org2"
SETTING_MINTRUD_ORG2_NAME = "mintrud_org2_name"
SETTING_MINTRUD_TRAINED_REGISTRY_XLSX = "mintrud_trained_registry_xlsx_path"


def load_mintrud_employer_from_db() -> tuple[str, str, str, str]:
    """ИНН/наименование работодателя и организации 2 для шаблона Excel Минтруда."""
    return (
        _app_setting_get(SETTING_MINTRUD_INN_EMPLOYER, ""),
        _app_setting_get(SETTING_MINTRUD_EMPLOYER_NAME, ""),
        _app_setting_get(SETTING_MINTRUD_INN_ORG2, ""),
        _app_setting_get(SETTING_MINTRUD_ORG2_NAME, ""),
    )


def load_mintrud_trained_registry_path() -> str:
    return _app_setting_get(SETTING_MINTRUD_TRAINED_REGISTRY_XLSX, "")


def save_mintrud_trained_registry_path(path: str) -> None:
    _app_setting_set(SETTING_MINTRUD_TRAINED_REGISTRY_XLSX, (path or "").strip())


def save_mintrud_employer_to_db(
    inn_employer: str,
    employer_name: str,
    inn_org2: str = "",
    org2_name: str = "",
) -> None:
    _app_setting_set(SETTING_MINTRUD_INN_EMPLOYER, (inn_employer or "").strip())
    _app_setting_set(SETTING_MINTRUD_EMPLOYER_NAME, (employer_name or "").strip())
    _app_setting_set(SETTING_MINTRUD_INN_ORG2, (inn_org2 or "").strip())
    _app_setting_set(SETTING_MINTRUD_ORG2_NAME, (org2_name or "").strip())


def build_commission_template_payload(
    format_date_words: Callable[[str], str],
    *,
    kind: str = COMMISSION_KIND_OT,
) -> dict[str, str]:
    """
    Данные из блока «Приказ и комиссия» (SQLite) для вставки в шаблон протокола.
    Дата — в том же словесном виде, что дата протокола; ФИО и должности — родительный падеж
    (при установленном pymorphy2, иначе без изменения).
    """
    order_no, order_date, chair, members = load_commission_state_from_db(kind)
    venue, order_approver = load_commission_protocol_context_from_db(kind)
    date_w = ""
    if (order_date or "").strip():
        date_w = format_date_words(order_date.strip())
    on = (order_no or "").strip()
    chair_s = ""
    if chair is not None:
        chair_s = format_person_fio_profession_genitive(chair.fio, chair.profession)
    mem_parts: list[str] = []
    for m in members:
        if not (m.fio or "").strip():
            continue
        mem_parts.append(format_person_fio_profession_genitive(m.fio, m.profession))
    members_s = _format_commission_members_two_columns(mem_parts)
    return {
        "date_words": date_w,
        "order_no": on,
        "chair": chair_s,
        "members": members_s,
        "venue_subdivision": venue,
        "order_approver": order_approver,
    }


def _format_commission_members_two_columns(mem_parts: list[str]) -> str:
    """
    Члены комиссии в шапке протокола: два столбца в одной зоне (табуляция),
    чтобы список не занимал полстраницы по вертикали.
    """
    rows = commission_members_two_column_rows(mem_parts)
    if not rows:
        return ""
    if len(rows) == 1 and not rows[0][1]:
        return rows[0][0]
    lines: list[str] = []
    for left, right in rows:
        if left and right:
            lines.append(f"{left}\t{right}")
        elif left:
            lines.append(left)
        else:
            lines.append(right)
    return "\n".join(lines)


def commission_members_two_column_rows(mem_parts: list[str]) -> list[tuple[str, str]]:
    """Пары (левый столбец, правый) для шапки протокола."""
    if not mem_parts:
        return []
    if len(mem_parts) == 1:
        return [(mem_parts[0], "")]
    mid = (len(mem_parts) + 1) // 2
    left = mem_parts[:mid]
    right = mem_parts[mid:]
    rows: list[tuple[str, str]] = []
    for i in range(max(len(left), len(right))):
        left_txt = left[i] if i < len(left) else ""
        r = right[i] if i < len(right) else ""
        rows.append((left_txt, r))
    return rows


def parse_commission_members_two_column_text(members_text: str) -> list[tuple[str, str]]:
    """Разбор текста членов комиссии после apply_commission_insertions (строки и табы)."""
    rows: list[tuple[str, str]] = []
    for line in (members_text or "").split("\n"):
        s = line.strip()
        if not s:
            continue
        if "\t" in line:
            left, right = line.split("\t", 1)
            rows.append((left.strip(), right.strip()))
        else:
            rows.append((s, ""))
    return rows


def commission_members_anchor_prefix(line: str) -> str:
    """Текст до конца слова «членов» / «членов комиссии» включительно (для абзаца шапки)."""
    lo = (line or "").lower()
    for phrase in ("членов комиссии", "членов"):
        i = lo.find(phrase)
        if i != -1:
            return line[: i + len(phrase)]
    return line or ""


def commission_chair_anchor_prefix(line: str) -> str:
    """Текст до конца слова «председателя» включительно (для абзаца шапки)."""
    lo = (line or "").lower()
    for phrase in ("председателя комиссии", "председателя"):
        i = lo.find(phrase)
        if i != -1:
            return line[: i + len(phrase)]
    return line or ""


def build_commission_signature_suffix_payload(
    *,
    kind: str = COMMISSION_KIND_OT,
) -> tuple[str, str]:
    """
    Блоки подписей для плейсхолдеров в конце протокола (.docx): председатель и члены.

    Формат (каждый блок — две строки с мягким переносом в Word):
      Председатель комиссии\\t_____И.О. Фамилия___________________
      <отступ>должность, подпись  или  И.О. Фамилия, подпись
      Члены комиссии:\\t_____И.О. Фамилия___________________
      <отступ>…
    """
    _, _, chair, members = load_commission_state_from_db(kind)
    chair_s = ""
    if chair is not None and (chair.fio or "").strip():
        chair_s = _format_chair_signature_block(chair.fio, chair.profession)
    blocks: list[str] = []
    for m in members:
        if not (m.fio or "").strip():
            continue
        blocks.append(_format_member_signature_block(m.fio, m.profession))
    members_s = "\n\n".join(blocks)
    return chair_s, members_s


def _tail_already_has_block(tail_after_kw: str, block: str) -> bool:
    """Не дублировать вставку, если после ключевого слова уже стоит тот же текст (с переноса строки)."""
    if not block.strip():
        return False
    t = tail_after_kw.lstrip()
    first_line = block.split("\n", 1)[0].strip()
    if len(first_line) < 2:
        return False
    if t.startswith("\n"):
        t = t[1:].lstrip()
    return t.startswith(first_line[: min(16, len(first_line))])


def _line_qualifies_for_commission_fill(text: str) -> bool:
    lo = text.lower()
    return (
        "комисс" in lo
        or "председател" in lo
        or "членов" in lo
        or ("приказ" in lo and " от" in text)
    )


# Бланк с плейсхолдером: от «__» … 20__ г. (часто в протоколе по техническим вопросам)
_COMMISSION_DATE_PLACEHOLDER_RE = re.compile(
    r"(?P<sp>\s)от\s*«[^»]{0,40}»\s*[\s_\u00a0]*\s*20__\s*г\.",
    re.IGNORECASE | re.UNICODE,
)


def _replace_commission_date_placeholder(norm: str, date_words: str) -> tuple[str, bool]:
    dw = (date_words or "").strip()
    if not dw:
        return norm, False
    m = _COMMISSION_DATE_PLACEHOLDER_RE.search(norm)
    if not m:
        return norm, False
    out = norm[: m.start()] + m.group("sp") + "от " + dw + norm[m.end() :]
    return out, True


def _replace_commission_order_underscores(norm: str, order_no: str) -> str:
    """№ ___ / №___ в строке приказа → № <номер> (не трогаем «ПРОТОКОЛ №»)."""
    on = (order_no or "").strip()
    if not on:
        return norm
    for m in re.finditer(r"\s*№\s*_{2,}\s*", norm):
        j = m.start()
        prev = norm[max(0, j - 14) : j]
        if prev.rstrip().endswith("ПРОТОКОЛ") or prev.lower().rstrip().endswith("протокол"):
            continue
        return norm[: m.start()] + " № " + on + norm[m.end() :]
    return norm


def apply_commission_insertions_to_line(
    text: str,
    *,
    date_words: str,
    order_no: str,
    chair_gen: str,
    members_gen: str,
) -> str:
    """
    После « от» — дата приказа словами; после подходящего « №» (не «ПРОТОКОЛ №») — номер;
    после «председателя» — с новой строки; после «членов»/«членов комиссии» — в два столбца
    (табуляция, по половине списка), в каждой ячейке ФИО и должность (род. п.).
    Плейсхолдер «от «__» … 20__ г.» и «№ ___» заменяются целиком (как в бланках с подчёркиваниями).
    """
    if not _line_qualifies_for_commission_fill(text):
        return text
    base_norm = text.replace("\xa0", " ")
    norm = base_norm
    norm, date_placeholder_done = _replace_commission_date_placeholder(norm, date_words)
    norm = _replace_commission_order_underscores(norm, order_no)
    ops: list[tuple[int, str]] = []

    if date_words and not date_placeholder_done:
        key = " от"
        idx = norm.find(key)
        if idx != -1:
            pos = idx + len(key)
            rest = norm[pos:].lstrip()
            if not rest.startswith("«"):
                ops.append((pos, " " + date_words))

    if order_no:
        search_from = 0
        while True:
            j = norm.find(" №", search_from)
            if j == -1:
                break
            window = norm[max(0, j - 12) : j]
            if window.rstrip().endswith("ПРОТОКОЛ"):
                search_from = j + 1
                continue
            pos = j + len(" №")
            rest = norm[pos:].lstrip()
            if rest.startswith(order_no):
                break
            ops.append((pos, " " + order_no))
            break

    if chair_gen:
        lo = norm.lower()
        kw = "председателя"
        i = lo.find(kw)
        if i != -1:
            pos = i + len(kw)
            tail = norm[pos:]
            ins = "\n" + chair_gen
            if not _tail_already_has_block(tail, chair_gen):
                ops.append((pos, ins))

    if members_gen:
        lo = norm.lower()
        pos = -1
        for phrase in ("членов комиссии", "членов"):
            i = lo.find(phrase)
            if i != -1:
                pos = i + len(phrase)
                break
        if pos != -1:
            tail = norm[pos:]
            ins = "\n" + members_gen
            if not _tail_already_has_block(tail, members_gen):
                ops.append((pos, ins))

    if not ops:
        return text if norm == base_norm else norm
    ops.sort(key=lambda x: -x[0])
    out = norm
    for pos, ins in ops:
        out = out[:pos] + ins + out[pos:]
    return out


@dataclass
class CommissionState:
    """Кэш списка с листа komission и выбранный состав (синхронизируется с БД по кнопке «Сохранить»)."""

    pool: list[EmployeeRecord] = field(default_factory=list)
    chair: EmployeeRecord | None = None
    members: list[EmployeeRecord] = field(default_factory=list)


def refresh_commission_pool_from_excel(
    state: CommissionState,
    path: Path,
    *,
    show_errors: bool = True,
    parent: tk.Misc | None = None,
) -> None:
    try:
        state.pool = load_commission_from_excel(path)
        sort_employees_by_fio_alphabet(state.pool)
    except EmployeeExcelError as e:
        state.pool = []
        if show_errors:
            messagebox.showerror("Лист комиссии", str(e), parent=parent)


class CommissionAdminPanel(ttk.Labelframe):
    """Форма: № и дата приказа, выбор председателя и членов из state.pool."""

    def __init__(
        self,
        master: tk.Misc,
        *,
        state: CommissionState,
        get_excel_path: Callable[[], Path],
        dialog_parent: tk.Misc,
        commission_kind: str = COMMISSION_KIND_OT,
        mirror_pool_state: CommissionState | None = None,
    ) -> None:
        lf_title = (
            "Приказ и комиссия (проверка технических знаний)"
            if commission_kind == COMMISSION_KIND_TECH
            else "Приказ и комиссия по проверке знаний работников"
        )
        super().__init__(master, text=lf_title, padding=8, style="Card.TLabelframe")
        self._state = state
        self._get_excel_path = get_excel_path
        self._dialog_parent = dialog_parent
        self._commission_kind = commission_kind
        self._mirror_pool_state = mirror_pool_state
        g = pad()
        self.columnconfigure(1, weight=1)
        R = 3

        ttk.Label(
            self,
            text="Сохранённая комиссия (название / подразделение):",
        ).grid(row=0, column=0, sticky=tk.W, **g)
        prof_fr = ttk.Frame(self)
        prof_fr.grid(row=0, column=1, sticky=tk.EW, **g)
        prof_fr.columnconfigure(0, weight=1)
        self.var_commission_profile = tk.StringVar(value="")
        self.cmb_commission_profile = ttk.Combobox(
            prof_fr,
            textvariable=self.var_commission_profile,
            width=46,
            style=FIELD_COMBO_STYLE,
        )
        self.cmb_commission_profile.grid(row=0, column=0, sticky=tk.EW, padx=(0, 6))
        self.cmb_commission_profile.bind("<<ComboboxSelected>>", self._on_profile_selected)
        prof_btns = ttk.Frame(self)
        prof_btns.grid(row=1, column=0, columnspan=2, sticky=tk.W, padx=g["padx"])
        ttk.Button(prof_btns, text="Загрузить", command=self._load_selected_profile).grid(
            row=0, column=0, padx=(0, 6)
        )
        ttk.Button(prof_btns, text="Удалить профиль", command=self._delete_profile).grid(
            row=0, column=1, padx=(0, 6)
        )
        ttk.Label(
            self,
            text=(
                "Пример названия: «НПС и ЦТТ». Сохранение ниже обновляет выбранный профиль "
                "и подставляет его в протокол."
            ),
            wraplength=480,
            style="Hint.TLabel",
        ).grid(row=2, column=0, columnspan=2, sticky=tk.W, padx=g["padx"], pady=(0, g["pady"]))

        ttk.Label(self, text="№ приказа о комиссии:").grid(row=R + 0, column=0, sticky=tk.W, **g)
        self.entry_commission_order_no = ttk.Entry(self, width=50, style=FIELD_STYLE)
        self.entry_commission_order_no.grid(row=R + 0, column=1, sticky=tk.EW, **g)

        ttk.Label(self, text="Дата приказа (ДД.ММ.ГГГГ):").grid(row=R + 1, column=0, sticky=tk.W, **g)
        self.entry_commission_order_date = ttk.Entry(self, width=50, style=FIELD_DATE_STYLE)
        self.entry_commission_order_date.grid(row=R + 1, column=1, sticky=tk.EW, **g)

        ttk.Label(
            self,
            text="Подразделение (место проверки знаний):",
        ).grid(row=R + 2, column=0, sticky=tk.NW, **g)
        self.txt_commission_venue = tk.Text(self, height=3, width=52, wrap=tk.WORD)
        configure_editable_text(self.txt_commission_venue)
        self.txt_commission_venue.grid(row=R + 2, column=1, sticky=tk.EW, **g)

        ttk.Label(
            self,
            text='В соответствии с приказом (кем утверждён, без этих слов):',
        ).grid(row=R + 3, column=0, sticky=tk.W, **g)
        self.entry_commission_order_approver = ttk.Entry(self, width=50, style=FIELD_STYLE)
        self.entry_commission_order_approver.grid(row=R + 3, column=1, sticky=tk.EW, **g)

        ttk.Label(
            self,
            text=(
                "Текст подставляется в шаблон Word по маркерам "
                f"{COMMISSION_VENUE_PLACEHOLDER} и {COMMISSION_ORDER_APPROVER_PLACEHOLDER} "
                "(не удаляйте их в .docx). Отдельно для вкладок «Охрана труда» и «Технич. вопросы»."
            ),
            wraplength=480,
            style="Hint.TLabel",
        ).grid(row=R + 4, column=0, columnspan=2, sticky=tk.W, padx=g["padx"], pady=(2, g["pady"]))

        ttk.Label(
            self,
            text=(
                "Кандидаты — лист «komission»: строка 2 — подписи «ФИО»/«Должность»; с 3-й строки — "
                "председатель в A и B, члены в D и E (несколько строк членов — только D/E). "
                "В списке — «ФИО — должность», без повторов."
            ),
            wraplength=480,
        ).grid(row=R + 5, column=0, columnspan=2, sticky=tk.W, padx=g["padx"], pady=(8, g["pady"]))

        pool_fr = ttk.Frame(self)
        pool_fr.grid(row=R + 6, column=0, columnspan=2, sticky=tk.NSEW, pady=(4, 0))
        pool_fr.columnconfigure(0, weight=1)
        sb_pool = ttk.Scrollbar(pool_fr)
        sb_pool.grid(row=0, column=1, sticky=tk.NS)
        self.list_commission_pool = tk.Listbox(
            pool_fr,
            height=5,
            exportselection=False,
            yscrollcommand=sb_pool.set,
        )
        configure_listbox(self.list_commission_pool)
        self.list_commission_pool.grid(row=0, column=0, sticky=tk.NSEW)
        sb_pool.configure(command=self.list_commission_pool.yview)

        pool_btns = ttk.Frame(self)
        pool_btns.grid(row=R + 7, column=0, columnspan=2, sticky=tk.W, pady=(6, 0))
        ttk.Button(
            pool_btns,
            text="Обновить из Excel",
            command=self._on_refresh_excel_clicked,
        ).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(pool_btns, text="Назначить председателем", command=self._set_chair).grid(
            row=0, column=1, padx=(0, 6)
        )
        ttk.Button(pool_btns, text="Добавить в члены комиссии", command=self._add_member).grid(
            row=0, column=2, padx=(0, 6)
        )

        ttk.Label(self, text="Председатель комиссии:").grid(
            row=R + 8, column=0, sticky=tk.NW, padx=g["padx"], pady=(10, g["pady"])
        )
        ch_fr = ttk.Frame(self)
        ch_fr.grid(row=R + 8, column=1, sticky=tk.EW, padx=g["padx"], pady=(10, g["pady"]))
        self.lbl_commission_chair = ttk.Label(ch_fr, text="— не выбран —", wraplength=420)
        self.lbl_commission_chair.grid(row=0, column=0, sticky=tk.W)
        ttk.Button(ch_fr, text="Сбросить", command=self._clear_chair, width=10).grid(
            row=0, column=1, sticky=tk.E, padx=(8, 0)
        )

        ttk.Label(self, text="Члены комиссии:").grid(row=R + 9, column=0, sticky=tk.NW, **g)
        mem_fr = ttk.Frame(self)
        mem_fr.grid(row=R + 9, column=1, sticky=tk.EW)
        mem_fr.columnconfigure(0, weight=1)
        sb_mem = ttk.Scrollbar(mem_fr)
        sb_mem.grid(row=0, column=1, sticky=tk.NS)
        self.list_commission_members = tk.Listbox(
            mem_fr,
            height=4,
            exportselection=False,
            yscrollcommand=sb_mem.set,
        )
        configure_listbox(self.list_commission_members)
        self.list_commission_members.grid(row=0, column=0, sticky=tk.NSEW)
        sb_mem.configure(command=self.list_commission_members.yview)
        mem_btns = ttk.Frame(self)
        mem_btns.grid(row=R + 10, column=1, sticky=tk.W, pady=(4, 0))
        ttk.Button(mem_btns, text="Удалить выбранного из членов", command=self._remove_member).grid(
            row=0, column=0, padx=(0, 8)
        )
        ttk.Button(mem_btns, text="Очистить список членов", command=self._clear_members).grid(
            row=0, column=1
        )

        ttk.Button(
            self,
            text="Сохранить приказ и состав в базу данных",
            command=self._save_to_db,
        ).grid(row=R + 11, column=0, columnspan=2, sticky=tk.W, pady=(12, 0))

        for w in (
            self.entry_commission_order_no,
            self.entry_commission_order_date,
            self.entry_commission_order_approver,
            self.txt_commission_venue,
        ):
            bind_editable_clipboard(w)
        try:
            self.after_idle(
                lambda: register_clipboard_window(self.winfo_toplevel())
            )
        except tk.TclError:
            pass

        self._refresh_profile_combobox()
        active = get_active_commission_profile_name(self._commission_kind)
        if active and load_commission_profile(active, self._commission_kind):
            self.var_commission_profile.set(active)
            self._apply_profile_to_ui(active, show_message=False)
        else:
            self.load_from_db_into_ui()
        self.refresh_pool_display()

    def _refresh_profile_combobox(self) -> None:
        names = list_commission_profile_names(self._commission_kind)
        self.cmb_commission_profile["values"] = names

    def _apply_profile_to_ui(self, name: str, *, show_message: bool = True) -> None:
        prof = load_commission_profile(name, self._commission_kind)
        if prof is None:
            if show_message:
                messagebox.showinfo(
                    "Комиссия",
                    f"Профиль «{name}» не найден.",
                    parent=self._dialog_parent,
                )
            return
        self.entry_commission_order_no.delete(0, tk.END)
        self.entry_commission_order_no.insert(0, prof["order_no"])
        self.entry_commission_order_date.delete(0, tk.END)
        self.entry_commission_order_date.insert(0, prof["order_date"])
        self.txt_commission_venue.delete("1.0", tk.END)
        if prof["venue_subdivision"]:
            self.txt_commission_venue.insert("1.0", prof["venue_subdivision"])
        self.entry_commission_order_approver.delete(0, tk.END)
        self.entry_commission_order_approver.insert(0, prof["order_approver"])
        self._state.chair = prof["chair"]
        self._state.members = list(prof["members"])
        self._refresh_chair_label()
        self._refresh_members_listbox()
        try:
            save_commission_state_to_db(
                prof["order_no"],
                prof["order_date"],
                prof["chair"],
                prof["members"],
                kind=self._commission_kind,
                venue_subdivision=prof["venue_subdivision"],
                order_approver=prof["order_approver"],
            )
        except sqlite3.Error as e:
            if show_message:
                messagebox.showerror("База данных", str(e), parent=self._dialog_parent)
            return
        set_active_commission_profile_name(name, kind=self._commission_kind)
        self.var_commission_profile.set(name)

    def _on_profile_selected(self, _evt: object | None = None) -> None:
        name = self.var_commission_profile.get().strip()
        if name:
            self._apply_profile_to_ui(name, show_message=False)

    def _load_selected_profile(self) -> None:
        name = self.var_commission_profile.get().strip()
        if not name:
            messagebox.showinfo(
                "Комиссия",
                "Выберите или введите название сохранённой комиссии.",
                parent=self._dialog_parent,
            )
            return
        self._apply_profile_to_ui(name)

    def _delete_profile(self) -> None:
        name = self.var_commission_profile.get().strip()
        if not name:
            messagebox.showinfo(
                "Комиссия",
                "Выберите профиль для удаления.",
                parent=self._dialog_parent,
            )
            return
        if not messagebox.askyesno(
            "Комиссия",
            f"Удалить сохранённый профиль «{name}»?",
            parent=self._dialog_parent,
        ):
            return
        try:
            delete_commission_profile(name, self._commission_kind)
        except sqlite3.Error as e:
            messagebox.showerror("База данных", str(e), parent=self._dialog_parent)
            return
        self.var_commission_profile.set("")
        self._refresh_profile_combobox()

    def _on_refresh_excel_clicked(self) -> None:
        refresh_commission_pool_from_excel(
            self._state,
            self._get_excel_path(),
            show_errors=True,
            parent=self._dialog_parent,
        )
        if self._mirror_pool_state is not None:
            self._mirror_pool_state.pool = self._state.pool
        self.refresh_pool_display()

    def refresh_pool_display(self) -> None:
        self.list_commission_pool.delete(0, tk.END)
        for rec in self._state.pool:
            self.list_commission_pool.insert(tk.END, listbox_label_for_employee(rec))

    def load_from_db_into_ui(self) -> None:
        on, od, chair, members = load_commission_state_from_db(self._commission_kind)
        venue, approver = load_commission_protocol_context_from_db(self._commission_kind)
        self.entry_commission_order_no.delete(0, tk.END)
        self.entry_commission_order_no.insert(0, on)
        self.entry_commission_order_date.delete(0, tk.END)
        self.entry_commission_order_date.insert(0, od)
        self.txt_commission_venue.delete("1.0", tk.END)
        if venue:
            self.txt_commission_venue.insert("1.0", venue)
        self.entry_commission_order_approver.delete(0, tk.END)
        self.entry_commission_order_approver.insert(0, approver)
        self._state.chair = chair
        self._state.members = list(members)
        self._refresh_chair_label()
        self._refresh_members_listbox()

    def _refresh_chair_label(self) -> None:
        if self._state.chair is None:
            self.lbl_commission_chair.configure(text="— не выбран —")
        else:
            self.lbl_commission_chair.configure(text=listbox_label_for_employee(self._state.chair))

    def _refresh_members_listbox(self) -> None:
        self.list_commission_members.delete(0, tk.END)
        for m in self._state.members:
            self.list_commission_members.insert(tk.END, listbox_label_for_employee(m))

    def _strip_members_equal_chair(self) -> None:
        if self._state.chair is None:
            return
        k_ch = employee_unique_key(self._state.chair)
        self._state.members = [m for m in self._state.members if employee_unique_key(m) != k_ch]
        self._refresh_members_listbox()

    def _set_chair(self) -> None:
        sel = self.list_commission_pool.curselection()
        if not sel:
            messagebox.showinfo(
                "Комиссия",
                "Выберите строку в списке кандидатов.",
                parent=self._dialog_parent,
            )
            return
        idx = int(sel[0])
        if not (0 <= idx < len(self._state.pool)):
            return
        self._state.chair = self._state.pool[idx]
        self._refresh_chair_label()
        self._strip_members_equal_chair()

    def _add_member(self) -> None:
        sel = self.list_commission_pool.curselection()
        if not sel:
            messagebox.showinfo(
                "Комиссия",
                "Выберите строку в списке кандидатов.",
                parent=self._dialog_parent,
            )
            return
        idx = int(sel[0])
        cand = self._state.pool[idx]
        k = employee_unique_key(cand)
        if self._state.chair and employee_unique_key(self._state.chair) == k:
            messagebox.showinfo(
                "Комиссия",
                "Этот человек уже назначен председателем.",
                parent=self._dialog_parent,
            )
            return
        if any(employee_unique_key(m) == k for m in self._state.members):
            messagebox.showinfo(
                "Комиссия",
                "Этот человек уже в списке членов комиссии.",
                parent=self._dialog_parent,
            )
            return
        self._state.members.append(cand)
        self._refresh_members_listbox()

    def _remove_member(self) -> None:
        sel = self.list_commission_members.curselection()
        if not sel:
            messagebox.showinfo(
                "Комиссия",
                "Выберите члена комиссии в списке ниже.",
                parent=self._dialog_parent,
            )
            return
        idx = int(sel[0])
        if 0 <= idx < len(self._state.members):
            del self._state.members[idx]
        self._refresh_members_listbox()

    def _clear_chair(self) -> None:
        self._state.chair = None
        self._refresh_chair_label()

    def _clear_members(self) -> None:
        self._state.members = []
        self._refresh_members_listbox()

    def _save_to_db(self) -> None:
        on = self.entry_commission_order_no.get().strip()
        od = self.entry_commission_order_date.get().strip()
        venue = self.txt_commission_venue.get("1.0", tk.END).strip()
        approver = self.entry_commission_order_approver.get().strip()
        profile_name = self.var_commission_profile.get().strip()
        if not profile_name:
            profile_name = venue.splitlines()[0].strip() if venue else ""
        if not profile_name:
            messagebox.showinfo(
                "Комиссия",
                "Введите название профиля (подразделение) в поле вверху, "
                "например «НПС и ЦТТ».",
                parent=self._dialog_parent,
            )
            return
        try:
            save_commission_profile(
                profile_name,
                self._commission_kind,
                order_no=on,
                order_date=od,
                chair=self._state.chair,
                members=self._state.members,
                venue_subdivision=venue,
                order_approver=approver,
            )
            save_commission_state_to_db(
                on,
                od,
                self._state.chair,
                self._state.members,
                kind=self._commission_kind,
                venue_subdivision=venue,
                order_approver=approver,
            )
        except sqlite3.Error as e:
            messagebox.showerror("База данных", str(e), parent=self._dialog_parent)
            return
        except ValueError as e:
            messagebox.showwarning("Комиссия", str(e), parent=self._dialog_parent)
            return
        self._refresh_profile_combobox()
        self.var_commission_profile.set(profile_name)
        msg = (
            f"Профиль «{profile_name}» и активная комиссия для протокола сохранены."
            if self._commission_kind != COMMISSION_KIND_TECH
            else f"Профиль «{profile_name}» (тех. вопросы) сохранён."
        )
        messagebox.showinfo("Комиссия", msg, parent=self._dialog_parent)
