# -*- coding: utf-8 -*-
"""Диалоги добавления сотрудника и восстановления из архива Excel."""

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk
from typing import Callable

from employees_io import EmployeeRecord, listbox_label_for_employee
from ui_theme import FIELD_STYLE, configure_listbox, pad


class EmployeeFormDialog:
    """Форма полей Data_base: ФИО, подразделение, должность, совмещение, СНИЛС."""

    _FIELDS: tuple[tuple[str, str, bool], ...] = (
        ("fio", "Фамилия, имя, отчество", True),
        ("subdivision", "Подразделение", False),
        ("profession", "Должность", False),
        ("profession2", "Совмещаемая профессия (если есть)", False),
        ("snils", "№ страхового свидетельства (СНИЛС)", False),
    )

    def __init__(
        self,
        parent: tk.Misc,
        *,
        title: str = "Добавить сотрудника",
        initial: EmployeeRecord | None = None,
        on_submit: Callable[[EmployeeRecord], None],
        themed_toplevel: Callable[[tk.Misc | None], tk.Toplevel],
        make_modal: Callable[[tk.Toplevel], None],
    ) -> None:
        self._on_submit = on_submit
        self._vars: dict[str, tk.StringVar] = {}
        win = themed_toplevel(parent)
        win.title(title)
        win.minsize(480, 320)
        outer = ttk.Frame(win, padding=10)
        outer.pack(fill=tk.BOTH, expand=True)
        outer.columnconfigure(1, weight=1)
        g = pad()
        for row, (key, label, required) in enumerate(self._FIELDS):
            req = " *" if required else ""
            ttk.Label(outer, text=f"{label}{req}:").grid(row=row, column=0, sticky=tk.NW, **g)
            var = tk.StringVar(
                value=getattr(initial, key, "") if initial is not None else ""
            )
            self._vars[key] = var
            entry = ttk.Entry(outer, textvariable=var, width=52, style=FIELD_STYLE)
            entry.grid(row=row, column=1, sticky=tk.EW, **g)
            if row == 0:
                entry.focus_set()
        ttk.Label(
            outer,
            text="Данные сохраняются в Data_base.xlsx (лист rabotnik). "
            "Если указана совмещаемая профессия — в файл добавляются две строки "
            "(основная должность и совмещение). "
            "Перед записью создаётся резервная копия «…_before_edit.xlsx».",
            wraplength=440,
            style="Muted.TLabel",
        ).grid(row=len(self._FIELDS), column=0, columnspan=2, sticky=tk.W, pady=(8, 0))
        btns = ttk.Frame(outer)
        btns.grid(row=len(self._FIELDS) + 1, column=0, columnspan=2, sticky=tk.E, pady=(12, 0))
        ttk.Button(btns, text="Сохранить", command=lambda: self._save(win)).grid(
            row=0, column=0, padx=(0, 8)
        )
        ttk.Button(btns, text="Отмена", command=win.destroy).grid(row=0, column=1)
        make_modal(win)

    def _save(self, win: tk.Toplevel) -> None:
        fio = self._vars["fio"].get().strip()
        if not fio:
            messagebox.showwarning("Сотрудник", "Укажите ФИО.", parent=win)
            return
        record = EmployeeRecord(
            fio=fio,
            subdivision=self._vars["subdivision"].get().strip(),
            profession=self._vars["profession"].get().strip(),
            profession2=self._vars["profession2"].get().strip(),
            snils=self._vars["snils"].get().strip(),
        )
        self._on_submit(record)
        win.destroy()


class EmployeeArchiveDialog:
    """Список сотрудников на листе rabotnik_archive с восстановлением."""

    def __init__(
        self,
        parent: tk.Misc,
        records: list[EmployeeRecord],
        *,
        on_restore: Callable[[list[EmployeeRecord]], bool | None],
        themed_toplevel: Callable[[tk.Misc | None], tk.Toplevel],
        make_modal: Callable[[tk.Toplevel], None],
    ) -> None:
        self._on_restore = on_restore
        win = themed_toplevel(parent)
        win.title("Архив сотрудников")
        win.minsize(560, 360)
        win.geometry("640x420")
        outer = ttk.Frame(win, padding=8)
        outer.pack(fill=tk.BOTH, expand=True)
        outer.rowconfigure(1, weight=1)
        outer.columnconfigure(0, weight=1)
        ttk.Label(
            outer,
            text=(
                "Сотрудники, перенесённые в архив (лист rabotnik_archive в Data_base.xlsx). "
                "Выберите одного или нескольких и нажмите «Восстановить»."
            ),
            wraplength=600,
            style="Muted.TLabel",
        ).grid(row=0, column=0, sticky=tk.W, pady=(0, 6))
        list_fr = ttk.Frame(outer)
        list_fr.grid(row=1, column=0, sticky=tk.NSEW)
        list_fr.rowconfigure(0, weight=1)
        list_fr.columnconfigure(0, weight=1)
        sb = ttk.Scrollbar(list_fr)
        lb = tk.Listbox(
            list_fr,
            selectmode=tk.EXTENDED,
            exportselection=False,
            yscrollcommand=sb.set,
        )
        configure_listbox(lb, mono=True)
        lb.grid(row=0, column=0, sticky=tk.NSEW)
        sb.grid(row=0, column=1, sticky=tk.NS)
        sb.configure(command=lb.yview)
        self._records = list(records)
        for rec in self._records:
            lb.insert(tk.END, listbox_label_for_employee(rec))
        if not self._records:
            lb.insert(tk.END, "(архив пуст)")
            lb.configure(state=tk.DISABLED)
        btns = ttk.Frame(outer)
        btns.grid(row=2, column=0, sticky=tk.E, pady=(8, 0))
        ttk.Button(
            btns,
            text="Восстановить",
            command=lambda: self._restore(win, lb),
            state=tk.NORMAL if self._records else tk.DISABLED,
        ).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(btns, text="Закрыть", command=win.destroy).grid(row=0, column=1)
        make_modal(win)

    def _restore(self, win: tk.Toplevel, lb: tk.Listbox) -> None:
        sel = lb.curselection()
        if not sel:
            messagebox.showinfo("Архив", "Выберите сотрудников в списке.", parent=win)
            return
        chosen = [self._records[int(i)] for i in sel]
        ok = self._on_restore(chosen)
        if ok is False:
            return
        win.destroy()
