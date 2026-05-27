# -*- coding: utf-8 -*-
"""Выбор совмещаемых должностей для программ «В» (несколько строк Excel с одним ФИО)."""

from __future__ import annotations

import tkinter as tk
from dataclasses import dataclass, field
from tkinter import messagebox, ttk

from employees_io import EmployeeRecord
from v_program_registry_match import norm_profession_key


@dataclass
class VProfCombinationConfig:
    """Настройки совмещений: отдельно по каждому ФИО."""

    main_by_fio: dict[str, str] = field(default_factory=dict)
    enabled_by_fio: dict[str, frozenset[str]] = field(default_factory=dict)

    def global_enabled_norm_keys(self) -> frozenset[str]:
        """Объединение отмеченных должностей (для шапки протокола)."""
        out: set[str] = set()
        for keys in self.enabled_by_fio.values():
            out.update(keys)
        return frozenset(out)


def professions_ordered_from_records(records: list[EmployeeRecord]) -> list[str]:
    """Уникальные должности из записей (основная и совмещаемая), порядок первого появления."""
    out: list[str] = []
    seen: set[str] = set()
    for rec in records:
        for pr in (rec.profession, rec.profession2):
            t = (pr or "").strip()
            if not t:
                continue
            k = norm_profession_key(t)
            if k in seen:
                continue
            seen.add(k)
            out.append(t)
    return out


def professions_by_fio(
    records: list[EmployeeRecord],
) -> list[tuple[str, str, list[str]]]:
    """
    Группы по ФИО: (ключ, отображаемое ФИО, список уникальных должностей).
    Порядок — как в records.
    """
    order: list[str] = []
    buckets: dict[str, list[str]] = {}
    display: dict[str, str] = {}
    for rec in records:
        fio = (rec.fio or "").strip()
        key = norm_profession_key(fio) or f"__empty_{len(order)}__"
        if key not in buckets:
            buckets[key] = []
            order.append(key)
            display[key] = fio or "(без ФИО)"
        profs = professions_ordered_from_records([rec])
        for p in profs:
            pk = norm_profession_key(p)
            if not any(norm_profession_key(x) == pk for x in buckets[key]):
                buckets[key].append(p)
    return [(k, display[k], buckets[k]) for k in order]


def needs_combinations_dialog(records: list[EmployeeRecord]) -> bool:
    """Нужен выбор совмещений, если у кого-то больше одной должности в выборе."""
    for _key, _fio, profs in professions_by_fio(records):
        if len(profs) > 1:
            return True
    return False


def selection_signature(records: list[EmployeeRecord]) -> str:
    parts: list[str] = []
    for key, fio, profs in professions_by_fio(records):
        parts.append(f"{key}:{','.join(norm_profession_key(p) for p in profs)}")
    return "|".join(parts)


class VProfCombinationsDialog(tk.Toplevel):
    """
    По каждому сотруднику (ФИО): галочки — учёт в «В», радиокнопка — основная должность в протоколе.
    """

    def __init__(
        self,
        master: tk.Misc,
        groups: list[tuple[str, str, list[str]]],
        *,
        initial_main_by_fio: dict[str, str] | None = None,
        initial_enabled_by_fio: dict[str, frozenset[str]] | None = None,
    ) -> None:
        super().__init__(master)
        self.title("Совмещения — программы «В»")
        self.transient(master)
        self.grab_set()
        self.resizable(True, True)
        self.minsize(520, 240)
        self._result: VProfCombinationConfig | None = None
        self._main_vars: dict[str, tk.StringVar] = {}
        self._check_vars: dict[tuple[str, str], tk.BooleanVar] = {}
        init_main = initial_main_by_fio or {}
        init_enabled = initial_enabled_by_fio or {}

        outer = ttk.Frame(self, padding=10)
        outer.pack(fill=tk.BOTH, expand=True)
        ttk.Label(
            outer,
            text=(
                "Для каждого сотрудника отметьте должности для программ «В» (шапка и таблица) "
                "и выберите основную — она будет в графе «Должность» протокола для этого человека."
            ),
            wraplength=480,
        ).pack(anchor=tk.W, pady=(0, 8))

        body = ttk.Frame(outer)
        body.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(body, highlightthickness=0)
        sb = ttk.Scrollbar(body, orient=tk.VERTICAL, command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind(
            "<Configure>",
            lambda _e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=inner, anchor=tk.NW)
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        for fio_key, fio, profs in groups:
            lf = ttk.Labelframe(inner, text=fio, padding=6)
            lf.pack(fill=tk.X, pady=(0, 8))
            enabled_set = init_enabled.get(fio_key)
            main_init = (init_main.get(fio_key) or "").strip()
            if not main_init and profs:
                main_init = profs[0]
            main_var = tk.StringVar(value=main_init)
            self._main_vars[fio_key] = main_var
            for pr in profs:
                pk = norm_profession_key(pr)
                enabled = True
                if enabled_set is not None:
                    enabled = pk in enabled_set
                chk_var = tk.BooleanVar(value=enabled)
                self._check_vars[(fio_key, pk)] = chk_var
                row = ttk.Frame(lf)
                row.pack(fill=tk.X, anchor=tk.W)
                ttk.Checkbutton(row, text=pr, variable=chk_var).pack(side=tk.LEFT)
                ttk.Radiobutton(
                    row,
                    text="основная",
                    value=pr,
                    variable=main_var,
                ).pack(side=tk.LEFT, padx=(12, 0))

        btns = ttk.Frame(outer)
        btns.pack(fill=tk.X, pady=(12, 0))
        ttk.Button(btns, text="OK", command=self._on_ok, width=10).pack(side=tk.RIGHT)
        ttk.Button(btns, text="Отмена", command=self._on_cancel, width=10).pack(
            side=tk.RIGHT, padx=(0, 8)
        )
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self.bind("<Escape>", lambda _e: self._on_cancel())
        self.update_idletasks()
        self.geometry(f"+{master.winfo_rootx() + 40}+{master.winfo_rooty() + 40}")

    def _on_ok(self) -> None:
        main_by_fio: dict[str, str] = {}
        enabled_by_fio: dict[str, frozenset[str]] = {}
        for fio_key, main_var in self._main_vars.items():
            enabled_keys: set[str] = set()
            for (fk, pk), var in self._check_vars.items():
                if fk != fio_key or not var.get():
                    continue
                enabled_keys.add(pk)
            if not enabled_keys:
                messagebox.showwarning(
                    "Совмещения",
                    "У каждого сотрудника отметьте хотя бы одну должность для программ «В».",
                    parent=self,
                )
                return
            main = main_var.get().strip()
            if not main or norm_profession_key(main) not in enabled_keys:
                messagebox.showwarning(
                    "Совмещения",
                    "Для каждого сотрудника выберите основную должность среди отмеченных.",
                    parent=self,
                )
                return
            main_by_fio[fio_key] = main
            enabled_by_fio[fio_key] = frozenset(enabled_keys)
        self._result = VProfCombinationConfig(
            main_by_fio=main_by_fio,
            enabled_by_fio=enabled_by_fio,
        )
        self.grab_release()
        self.destroy()

    def _on_cancel(self) -> None:
        self._result = None
        self.grab_release()
        self.destroy()

    def run(self) -> VProfCombinationConfig | None:
        self.wait_window()
        return self._result
