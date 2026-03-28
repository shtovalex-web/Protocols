# -*- coding: utf-8 -*-
"""Приказ о комиссии и состав комиссии: SQLite (app_settings) и блок интерфейса."""

from __future__ import annotations

import json
import sqlite3
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Callable

from russian_genitive import format_person_fio_profession_genitive

import tkinter as tk
from tkinter import messagebox
from tkinter import ttk

from employees_io import (
    EmployeeExcelError,
    EmployeeRecord,
    employee_unique_key,
    listbox_label_for_employee,
    load_commission_from_excel,
)

DATABASE_FILENAME = "protocols.db"

SETTING_COMMISSION_ORDER_NO = "commission_order_no"
SETTING_COMMISSION_ORDER_DATE = "commission_order_date"
SETTING_COMMISSION_CHAIR_JSON = "commission_chair_json"
SETTING_COMMISSION_MEMBERS_JSON = "commission_members_json"


def database_path() -> Path:
    return Path(__file__).resolve().parent / DATABASE_FILENAME


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
    )


def load_commission_state_from_db() -> tuple[str, str, EmployeeRecord | None, list[EmployeeRecord]]:
    order_no = _app_setting_get(SETTING_COMMISSION_ORDER_NO, "")
    order_date = _app_setting_get(SETTING_COMMISSION_ORDER_DATE, "")
    chair: EmployeeRecord | None = None
    raw_ch = _app_setting_get(SETTING_COMMISSION_CHAIR_JSON, "")
    if raw_ch.strip():
        try:
            c = _employee_from_settings_dict(json.loads(raw_ch))
            if c.fio.strip():
                chair = c
        except (json.JSONDecodeError, TypeError, ValueError):
            chair = None
    members: list[EmployeeRecord] = []
    raw_m = _app_setting_get(SETTING_COMMISSION_MEMBERS_JSON, "")
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


def save_commission_state_to_db(
    order_no: str,
    order_date: str,
    chair: EmployeeRecord | None,
    members: list[EmployeeRecord],
) -> None:
    _app_setting_set(SETTING_COMMISSION_ORDER_NO, (order_no or "").strip())
    _app_setting_set(SETTING_COMMISSION_ORDER_DATE, (order_date or "").strip())
    if chair is not None and (chair.fio or "").strip():
        _app_setting_set(
            SETTING_COMMISSION_CHAIR_JSON,
            json.dumps(asdict(chair), ensure_ascii=False),
        )
    else:
        _app_setting_set(SETTING_COMMISSION_CHAIR_JSON, "")
    _app_setting_set(
        SETTING_COMMISSION_MEMBERS_JSON,
        json.dumps(
            [asdict(m) for m in members if (m.fio or "").strip()],
            ensure_ascii=False,
        ),
    )


def build_commission_template_payload(
    format_date_words: Callable[[str], str],
) -> dict[str, str]:
    """
    Данные из блока «Приказ и комиссия» (SQLite) для вставки в шаблон протокола.
    Дата — в том же словесном виде, что дата протокола; ФИО и должности — родительный падеж
    (при установленном pymorphy2, иначе без изменения).
    """
    order_no, order_date, chair, members = load_commission_state_from_db()
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
    members_s = "\n".join(mem_parts)
    return {
        "date_words": date_w,
        "order_no": on,
        "chair": chair_s,
        "members": members_s,
    }


def build_commission_signature_suffix_payload() -> tuple[str, str]:
    """
    Председатель и члены для файла подписей в конце протокола — формат И.О. Фамилия
    (инициалы и фамилия; должность через запятую, как в Excel).
    """
    from employees_io import format_person_iof_line

    _, _, chair, members = load_commission_state_from_db()
    chair_s = ""
    if chair is not None and (chair.fio or "").strip():
        chair_s = format_person_iof_line(chair.fio, chair.profession)
    lines: list[str] = []
    for m in members:
        if not (m.fio or "").strip():
            continue
        lines.append(format_person_iof_line(m.fio, m.profession))
    return chair_s, "\n".join(lines)


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
    после «председателя» и после «членов»/«членов комиссии» — с новой строки столбиком ФИО и должность (род. п.).
    """
    if not _line_qualifies_for_commission_fill(text):
        return text
    norm = text.replace("\xa0", " ")
    ops: list[tuple[int, str]] = []

    if date_words:
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
        return text
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
    ) -> None:
        super().__init__(
            master,
            text="Приказ и комиссия по проверке знаний работников",
            padding=6,
        )
        self._state = state
        self._get_excel_path = get_excel_path
        self._dialog_parent = dialog_parent
        g = {"padx": 5, "pady": 5}
        self.columnconfigure(1, weight=1)

        ttk.Label(self, text="№ приказа о комиссии:").grid(row=0, column=0, sticky=tk.W, **g)
        self.entry_commission_order_no = ttk.Entry(self, width=50)
        self.entry_commission_order_no.grid(row=0, column=1, sticky=tk.EW, **g)

        ttk.Label(self, text="Дата приказа (ДД.ММ.ГГГГ):").grid(row=1, column=0, sticky=tk.W, **g)
        self.entry_commission_order_date = ttk.Entry(self, width=50)
        self.entry_commission_order_date.grid(row=1, column=1, sticky=tk.EW, **g)

        ttk.Label(
            self,
            text=(
                "Кандидаты — лист «komission»: с 3-й строки таблицы A+должность в B (председатель), "
                "D+должность в E (члены); в списке — «ФИО — должность», без повторов."
            ),
            wraplength=480,
        ).grid(row=2, column=0, columnspan=2, sticky=tk.W, padx=g["padx"], pady=(8, g["pady"]))

        pool_fr = ttk.Frame(self)
        pool_fr.grid(row=3, column=0, columnspan=2, sticky=tk.NSEW, pady=(4, 0))
        pool_fr.columnconfigure(0, weight=1)
        sb_pool = ttk.Scrollbar(pool_fr)
        sb_pool.grid(row=0, column=1, sticky=tk.NS)
        self.list_commission_pool = tk.Listbox(
            pool_fr,
            height=5,
            exportselection=False,
            font=("Segoe UI", 10),
            yscrollcommand=sb_pool.set,
        )
        self.list_commission_pool.grid(row=0, column=0, sticky=tk.NSEW)
        sb_pool.configure(command=self.list_commission_pool.yview)

        pool_btns = ttk.Frame(self)
        pool_btns.grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=(6, 0))
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
            row=5, column=0, sticky=tk.NW, padx=g["padx"], pady=(10, g["pady"])
        )
        ch_fr = ttk.Frame(self)
        ch_fr.grid(row=5, column=1, sticky=tk.EW, padx=g["padx"], pady=(10, g["pady"]))
        self.lbl_commission_chair = ttk.Label(ch_fr, text="— не выбран —", wraplength=420)
        self.lbl_commission_chair.grid(row=0, column=0, sticky=tk.W)
        ttk.Button(ch_fr, text="Сбросить", command=self._clear_chair, width=10).grid(
            row=0, column=1, sticky=tk.E, padx=(8, 0)
        )

        ttk.Label(self, text="Члены комиссии:").grid(row=6, column=0, sticky=tk.NW, **g)
        mem_fr = ttk.Frame(self)
        mem_fr.grid(row=6, column=1, sticky=tk.EW)
        mem_fr.columnconfigure(0, weight=1)
        sb_mem = ttk.Scrollbar(mem_fr)
        sb_mem.grid(row=0, column=1, sticky=tk.NS)
        self.list_commission_members = tk.Listbox(
            mem_fr,
            height=4,
            exportselection=False,
            font=("Segoe UI", 10),
            yscrollcommand=sb_mem.set,
        )
        self.list_commission_members.grid(row=0, column=0, sticky=tk.NSEW)
        sb_mem.configure(command=self.list_commission_members.yview)
        mem_btns = ttk.Frame(self)
        mem_btns.grid(row=7, column=1, sticky=tk.W, pady=(4, 0))
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
        ).grid(row=8, column=0, columnspan=2, sticky=tk.W, pady=(12, 0))

        self.load_from_db_into_ui()
        self.refresh_pool_display()

    def _on_refresh_excel_clicked(self) -> None:
        refresh_commission_pool_from_excel(
            self._state,
            self._get_excel_path(),
            show_errors=True,
            parent=self._dialog_parent,
        )
        self.refresh_pool_display()

    def refresh_pool_display(self) -> None:
        self.list_commission_pool.delete(0, tk.END)
        for rec in self._state.pool:
            self.list_commission_pool.insert(tk.END, listbox_label_for_employee(rec))

    def load_from_db_into_ui(self) -> None:
        on, od, chair, members = load_commission_state_from_db()
        self.entry_commission_order_no.delete(0, tk.END)
        self.entry_commission_order_no.insert(0, on)
        self.entry_commission_order_date.delete(0, tk.END)
        self.entry_commission_order_date.insert(0, od)
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
        try:
            save_commission_state_to_db(on, od, self._state.chair, self._state.members)
        except sqlite3.Error as e:
            messagebox.showerror("База данных", str(e), parent=self._dialog_parent)
            return
        messagebox.showinfo(
            "Комиссия",
            "Приказ и состав комиссии сохранены в базе данных.",
            parent=self._dialog_parent,
        )
