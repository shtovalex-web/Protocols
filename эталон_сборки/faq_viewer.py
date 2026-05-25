# -*- coding: utf-8 -*-
"""Окно справки: загрузка FAQ (FAQ.txt / FAQ.md) и поиск по тексту."""

from __future__ import annotations

from pathlib import Path

import tkinter as tk
from tkinter import messagebox, ttk

from app_paths import application_bundle_dir
from protocol_embedded_assets import apply_embedded_window_icon_from_parent

FAQ_CANDIDATE_NAMES = ("FAQ.txt", "FAQ.md")


def faq_file_path() -> Path:
    base = application_bundle_dir()
    for name in FAQ_CANDIDATE_NAMES:
        p = base / name
        if p.is_file():
            return p
    return base / FAQ_CANDIDATE_NAMES[0]


def open_faq_window(parent: tk.Misc, *, title: str = "Справка и FAQ") -> None:
    path = faq_file_path()
    if not path.is_file():
        messagebox.showerror(
            title,
            "Не найден файл справки (FAQ.txt или FAQ.md) в папке комплекта программы:\n"
            f"{application_bundle_dir()}",
            parent=parent,
        )
        return
    try:
        body_text = path.read_text(encoding="utf-8")
    except OSError as e:
        messagebox.showerror(title, f"Не удалось прочитать справку:\n{e}", parent=parent)
        return

    win = tk.Toplevel(parent)
    apply_embedded_window_icon_from_parent(win, parent)
    win.title(title)
    win.minsize(520, 400)
    win.geometry("700x520")

    bar = ttk.Frame(win, padding=6)
    bar.pack(fill=tk.X)
    ttk.Label(bar, text="Поиск (от 2 симв.):").pack(side=tk.LEFT, padx=(0, 6))
    var_q = tk.StringVar()
    ent = ttk.Entry(bar, textvariable=var_q, width=36)
    ent.pack(side=tk.LEFT, padx=(0, 6))
    status = ttk.Label(bar, text="")
    status.pack(side=tk.LEFT, padx=(8, 0))

    mid = ttk.Frame(win)
    mid.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))
    sb = ttk.Scrollbar(mid)
    sb.pack(side=tk.RIGHT, fill=tk.Y)
    txt = tk.Text(
        mid,
        wrap=tk.WORD,
        yscrollcommand=sb.set,
        font=("Segoe UI", 10),
        padx=8,
        pady=8,
    )
    txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    sb.config(command=txt.yview)

    txt.insert("1.0", body_text)
    txt.configure(state=tk.DISABLED)
    txt.tag_configure("hit", background="#fff3a0")

    def _all_enabled() -> None:
        txt.configure(state=tk.NORMAL)

    def _all_disabled() -> None:
        txt.configure(state=tk.DISABLED)

    def clear_hits() -> None:
        _all_enabled()
        txt.tag_remove("hit", "1.0", tk.END)
        _all_disabled()

    next_from: list[str] = ["1.0"]

    def highlight_all() -> None:
        clear_hits()
        q = var_q.get().strip()
        if len(q) < 2:
            status.configure(text="Введите не менее 2 символов")
            return
        count = 0
        _all_enabled()
        pos = "1.0"
        while True:
            idx = txt.search(q, pos, tk.END, nocase=True)
            if not idx:
                break
            end = f"{idx}+{len(q)}c"
            txt.tag_add("hit", idx, end)
            pos = end
            count += 1
        _all_disabled()
        next_from[0] = "1.0"
        status.configure(text=f"Найдено: {count}" if count else "Нет вхождений")
        if count:
            txt.see("1.0")

    def find_next() -> None:
        q = var_q.get().strip()
        if len(q) < 2:
            status.configure(text="Введите не менее 2 символов")
            return
        _all_enabled()
        idx = txt.search(q, next_from[0], tk.END, nocase=True)
        if not idx:
            idx = txt.search(q, "1.0", tk.END, nocase=True)
        if not idx:
            _all_disabled()
            status.configure(text="Не найдено")
            return
        end = f"{idx}+{len(q)}c"
        txt.tag_remove("sel", "1.0", tk.END)
        txt.tag_add("sel", idx, end)
        next_from[0] = end
        txt.mark_set(tk.INSERT, end)
        txt.see(idx)
        _all_disabled()
        status.configure(text="Следующее вхождение")

    ttk.Button(bar, text="Подсветить все", command=highlight_all).pack(
        side=tk.LEFT, padx=(0, 6)
    )
    ttk.Button(bar, text="Далее", command=find_next).pack(side=tk.LEFT)
    ent.bind("<Return>", lambda _e: find_next())

    bf = ttk.Frame(win, padding=(6, 0, 6, 6))
    bf.pack(fill=tk.X)
    ttk.Button(bf, text="Закрыть", command=win.destroy).pack(side=tk.RIGHT)

    try:
        win.transient(parent)
    except tk.TclError:
        pass
    ent.focus_set()
    win.lift()
