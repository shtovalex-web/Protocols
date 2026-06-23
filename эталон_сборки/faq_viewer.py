# -*- coding: utf-8 -*-
"""Окно справки: FAQ, журнал доработок и поиск по тексту."""

from __future__ import annotations

from pathlib import Path

import tkinter as tk
from tkinter import messagebox, ttk

from app_paths import application_bundle_dir
from clipboard_ui import register_clipboard_window
from protocol_embedded_assets import apply_embedded_window_icon_from_parent
from ui_theme import Colors, SPACING, apply_theme_to_window, configure_readonly_text
from ui_widgets import build_dialog_button_row, build_search_toolbar

FAQ_CANDIDATE_NAMES = ("FAQ.txt", "FAQ.md")
CHANGELOG_CANDIDATE_NAMES = ("ЖУРНАЛ_ДОРАБОТОК.md",)


def _first_existing_in_bundle(names: tuple[str, ...]) -> Path | None:
    base = application_bundle_dir()
    for name in names:
        p = base / name
        if p.is_file():
            return p
    root = Path(__file__).resolve().parent
    for name in names:
        p = root / "bundle" / name
        if p.is_file():
            return p
    return None


def faq_file_path() -> Path:
    p = _first_existing_in_bundle(FAQ_CANDIDATE_NAMES)
    if p is not None:
        return p
    return application_bundle_dir() / FAQ_CANDIDATE_NAMES[0]


def changelog_file_path() -> Path | None:
    return _first_existing_in_bundle(CHANGELOG_CANDIDATE_NAMES)


def open_text_help_window(
    parent: tk.Misc,
    path: Path,
    *,
    title: str,
    missing_hint: str = "",
) -> None:
    if not path.is_file():
        messagebox.showerror(
            title,
            missing_hint or f"Не найден файл:\n{path}",
            parent=parent,
        )
        return
    try:
        body_text = path.read_text(encoding="utf-8")
    except OSError as e:
        messagebox.showerror(title, f"Не удалось прочитать файл:\n{e}", parent=parent)
        return

    win = tk.Toplevel(parent)
    apply_theme_to_window(win)
    apply_embedded_window_icon_from_parent(win, parent)
    win.title(title)
    win.minsize(520, 400)
    win.geometry("720x540")

    outer = ttk.Frame(win, padding=SPACING)
    outer.pack(fill=tk.BOTH, expand=True)

    mid = ttk.Frame(outer)
    mid.pack(fill=tk.BOTH, expand=True, pady=(0, SPACING))
    sb = ttk.Scrollbar(mid)
    sb.pack(side=tk.RIGHT, fill=tk.Y)
    txt = tk.Text(mid, wrap=tk.WORD, yscrollcommand=sb.set)
    configure_readonly_text(txt)
    txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    sb.config(command=txt.yview)

    txt.insert("1.0", body_text)
    txt.configure(state=tk.DISABLED)
    txt.tag_configure("hit", background=Colors.search_hit)

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

    _bar, var_q, status, ent = build_search_toolbar(
        outer,
        on_highlight=highlight_all,
        on_next=find_next,
    )
    ent.bind("<Return>", lambda _e: find_next())

    build_dialog_button_row(outer, on_close=win.destroy)

    try:
        win.transient(parent)
    except tk.TclError:
        pass
    ent.focus_set()
    register_clipboard_window(win)
    win.lift()


def open_faq_window(parent: tk.Misc, *, title: str = "Справка и FAQ") -> None:
    path = faq_file_path()
    open_text_help_window(
        parent,
        path,
        title=title,
        missing_hint=(
            "Не найден файл справки (FAQ.txt или FAQ.md) в папке комплекта программы:\n"
            f"{application_bundle_dir()}"
        ),
    )


def open_changelog_window(parent: tk.Misc) -> None:
    path = changelog_file_path()
    if path is None:
        messagebox.showerror(
            "Журнал доработок",
            "Не найден файл ЖУРНАЛ_ДОРАБОТОК.md в папке комплекта программы:\n"
            f"{application_bundle_dir()}",
            parent=parent,
        )
        return
    open_text_help_window(parent, path, title="Журнал доработок по версиям")
