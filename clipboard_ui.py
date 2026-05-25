# -*- coding: utf-8 -*-
"""Вставка, копирование и вырезание из буфера обмена для полей tkinter (Windows)."""

from __future__ import annotations

import time
import tkinter as tk

_CLIPBOARD_INSTALLED = False
_PASTE_GUARD: dict[int, float] = {}
_PASTE_GUARD_SEC = 0.12

# Windows/ttk: вставка часто приходит и как Ctrl+V, и как <<Paste>> — одна операция.
_PASTE_SEQUENCES: tuple[str, ...] = (
    "<Control-v>",
    "<Control-V>",
    "<Control-Key-v>",
    "<Control-Key-V>",
    "<<Paste>>",
    "<Shift-Insert>",
)


def _clipboard_text(widget: tk.Misc) -> str:
    try:
        return widget.clipboard_get()
    except tk.TclError:
        return ""


def _delete_selection(widget: tk.Misc) -> None:
    try:
        if widget.selection_present():
            widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
    except tk.TclError:
        pass


def _widget_is_text_disabled(widget: tk.Misc) -> bool:
    try:
        return str(widget.cget("state")) == tk.DISABLED
    except tk.TclError:
        return False


def _paste_once(widget: tk.Misc, action) -> bool:
    """True — выполнить вставку; False — повтор того же нажатия (уже обработано)."""
    k = id(widget)
    now = time.monotonic()
    if now - _PASTE_GUARD.get(k, 0) < _PASTE_GUARD_SEC:
        return False
    _PASTE_GUARD[k] = now
    action()
    return True


def _paste_into_entry(widget: tk.Misc) -> None:
    text = _clipboard_text(widget)
    if not text:
        return
    _delete_selection(widget)
    widget.insert(tk.INSERT, text)


def _copy_from_entry(widget: tk.Misc) -> None:
    try:
        if widget.selection_present():
            widget.clipboard_clear()
            widget.clipboard_append(widget.selection_get())
    except tk.TclError:
        pass


def _cut_from_entry(widget: tk.Misc) -> None:
    _copy_from_entry(widget)
    _delete_selection(widget)


def _select_all_entry(widget: tk.Misc) -> None:
    try:
        widget.selection_range(0, tk.END)
        widget.icursor(tk.END)
    except tk.TclError:
        pass


def _paste_into_combobox(widget: tk.Misc) -> None:
    raw = _clipboard_text(widget)
    if not raw:
        return
    text = raw.replace("\r\n", "\n").split("\n", 1)[0].strip()
    if not text:
        return
    try:
        if str(widget.cget("state")) == "readonly":
            values = list(widget.cget("values") or ())
            if text in values:
                widget.set(text)
            return
    except tk.TclError:
        pass
    _delete_selection(widget)
    try:
        widget.insert(tk.INSERT, text)
    except tk.TclError:
        widget.set(((widget.get() or "") + text).strip())


def _paste_into_text(widget: tk.Misc) -> None:
    if _widget_is_text_disabled(widget):
        return
    text = _clipboard_text(widget)
    if not text:
        return
    _delete_selection(widget)
    widget.insert(tk.INSERT, text)


def _copy_from_text(widget: tk.Misc) -> None:
    if _widget_is_text_disabled(widget):
        return
    try:
        if widget.tag_ranges(tk.SEL):
            widget.clipboard_clear()
            widget.clipboard_append(widget.get(tk.SEL_FIRST, tk.SEL_LAST))
    except tk.TclError:
        pass


def _cut_from_text(widget: tk.Misc) -> None:
    if _widget_is_text_disabled(widget):
        return
    _copy_from_text(widget)
    _delete_selection(widget)


def _select_all_text(widget: tk.Misc) -> None:
    if _widget_is_text_disabled(widget):
        return
    try:
        widget.tag_add(tk.SEL, "1.0", tk.END)
        widget.mark_set(tk.INSERT, tk.END)
        widget.see(tk.INSERT)
    except tk.TclError:
        pass


def _make_paste_handler(paste_fn):
    def _handler(event: tk.Event) -> str:
        _paste_once(event.widget, lambda: paste_fn(event.widget))
        return "break"

    return _handler


def _on_entry_copy(event: tk.Event) -> str:
    _copy_from_entry(event.widget)
    return "break"


def _on_entry_cut(event: tk.Event) -> str:
    _cut_from_entry(event.widget)
    return "break"


def _on_entry_select_all(event: tk.Event) -> str:
    _select_all_entry(event.widget)
    return "break"


def _on_text_copy(event: tk.Event) -> str:
    _copy_from_text(event.widget)
    return "break"


def _on_text_cut(event: tk.Event) -> str:
    _cut_from_text(event.widget)
    return "break"


def _on_text_select_all(event: tk.Event) -> str:
    _select_all_text(event.widget)
    return "break"


def _popup_context_menu(event: tk.Event, *, widget_kind: str) -> str:
    w = event.widget
    menu = tk.Menu(w, tearoff=0)
    if widget_kind == "text" and _widget_is_text_disabled(w):
        try:
            if w.tag_ranges(tk.SEL):
                menu.add_command(
                    label="Копировать",
                    command=lambda: _copy_from_text(w),
                )
        except tk.TclError:
            pass
        if menu.index("end") is not None:
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()
        return "break"

    if widget_kind == "combobox":
        paste_fn = lambda: _paste_once(w, lambda: _paste_into_combobox(w))
        copy_fn = lambda: _copy_from_entry(w)
        cut_fn = lambda: _cut_from_entry(w)
        select_fn = lambda: _select_all_entry(w)
    elif widget_kind == "text":
        paste_fn = lambda: _paste_once(w, lambda: _paste_into_text(w))
        copy_fn = lambda: _copy_from_text(w)
        cut_fn = lambda: _cut_from_text(w)
        select_fn = lambda: _select_all_text(w)
    else:
        paste_fn = lambda: _paste_once(w, lambda: _paste_into_entry(w))
        copy_fn = lambda: _copy_from_entry(w)
        cut_fn = lambda: _cut_from_entry(w)
        select_fn = lambda: _select_all_entry(w)

    menu.add_command(label="Вырезать", command=cut_fn)
    menu.add_command(label="Копировать", command=copy_fn)
    menu.add_command(label="Вставить", command=paste_fn)
    menu.add_separator()
    menu.add_command(label="Выделить всё", command=select_fn)
    try:
        menu.tk_popup(event.x_root, event.y_root)
    finally:
        menu.grab_release()
    return "break"


def _bind_class_keys(
    root: tk.Misc,
    class_name: str,
    *,
    paste_handler,
    copy,
    cut,
    select_all,
    context_kind: str,
) -> None:
    for seq in _PASTE_SEQUENCES:
        root.bind_class(class_name, seq, paste_handler)
    for seq, handler in (
        ("<Control-c>", copy),
        ("<Control-C>", copy),
        ("<Control-Insert>", copy),
        ("<Control-x>", cut),
        ("<Control-X>", cut),
        ("<Shift-Delete>", cut),
        ("<Control-a>", select_all),
        ("<Control-A>", select_all),
        (
            "<Button-3>",
            lambda e, k=context_kind: _popup_context_menu(e, widget_kind=k),
        ),
    ):
        root.bind_class(class_name, seq, handler)


def _bind_global_paste_fallback(root: tk.Misc) -> None:
    """Запасной путь: фокусный виджет, если привязка класса не сработала."""

    def _route(event: tk.Event) -> str | None:
        w = root.focus_get()
        if w is None:
            return None
        cls = w.winfo_class()
        if cls in ("TEntry", "Entry"):
            if _paste_once(w, lambda: _paste_into_entry(w)):
                return "break"
        elif cls == "TCombobox":
            if _paste_once(w, lambda: _paste_into_combobox(w)):
                return "break"
        elif cls == "Text":
            if _paste_once(w, lambda: _paste_into_text(w)):
                return "break"
        return None

    for seq in _PASTE_SEQUENCES:
        root.bind_all(seq, _route, add="+")


def install_clipboard_support(root: tk.Misc) -> None:
    """
    Горячие клавиши и контекстное меню (ПКМ) для полей ввода во всём приложении.
    Вызывать один раз от корневого окна tk.Tk.
    """
    global _CLIPBOARD_INSTALLED
    if _CLIPBOARD_INSTALLED:
        return
    _CLIPBOARD_INSTALLED = True

    _bind_class_keys(
        root,
        "TEntry",
        paste_handler=_make_paste_handler(_paste_into_entry),
        copy=_on_entry_copy,
        cut=_on_entry_cut,
        select_all=_on_entry_select_all,
        context_kind="entry",
    )
    _bind_class_keys(
        root,
        "Entry",
        paste_handler=_make_paste_handler(_paste_into_entry),
        copy=_on_entry_copy,
        cut=_on_entry_cut,
        select_all=_on_entry_select_all,
        context_kind="entry",
    )
    _bind_class_keys(
        root,
        "TCombobox",
        paste_handler=_make_paste_handler(_paste_into_combobox),
        copy=_on_entry_copy,
        cut=_on_entry_cut,
        select_all=_on_entry_select_all,
        context_kind="combobox",
    )
    _bind_class_keys(
        root,
        "Text",
        paste_handler=_make_paste_handler(_paste_into_text),
        copy=_on_text_copy,
        cut=_on_text_cut,
        select_all=_on_text_select_all,
        context_kind="text",
    )
    _bind_global_paste_fallback(root)
