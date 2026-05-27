# -*- coding: utf-8 -*-
"""Вставка, копирование и вырезание из буфера обмена для полей tkinter (Windows)."""

from __future__ import annotations

import time
import tkinter as tk

_CLIPBOARD_INSTALLED = False
_PASTE_GUARD: dict[int, float] = {}
_PASTE_GUARD_SEC = 0.12
_DIRECT_BOUND: set[int] = set()

# Windows/ttk: вставка часто приходит и как Ctrl+V, и как <<Paste>> — одна операция.
_PASTE_SEQUENCES: tuple[str, ...] = (
    "<Control-v>",
    "<Control-V>",
    "<Control-Key-v>",
    "<Control-Key-V>",
    "<<Paste>>",
    "<Shift-Insert>",
)
# У Text свой <<Paste>>; переопределение ломает вставку на Windows.
_PASTE_SEQUENCES_TEXT: tuple[str, ...] = tuple(
    s for s in _PASTE_SEQUENCES if s != "<<Paste>>"
)


def _paste_sequences_for_class(class_name: str) -> tuple[str, ...]:
    if class_name == "Text":
        return _PASTE_SEQUENCES_TEXT
    return _PASTE_SEQUENCES


def _paste_sequences_for_widget(widget: tk.Misc) -> tuple[str, ...]:
    try:
        return _paste_sequences_for_class(widget.winfo_class())
    except tk.TclError:
        return _PASTE_SEQUENCES


def _clipboard_text(widget: tk.Misc) -> str:
    """Текст из буфера: на Windows иногда пусто через clipboard_get()."""
    root = widget.winfo_toplevel()
    sources: list[tk.Misc] = []
    for w in (widget, root):
        if w not in sources:
            sources.append(w)

    for w in sources:
        try:
            text = w.clipboard_get()
            if text:
                return str(text)
        except tk.TclError:
            pass
        for cmd in (
            ("tk::GetClipboardText",),
            ("clipboard", "get", "-format", "CF_UNICODETEXT"),
            ("clipboard", "get", "-format", "UnicodeText"),
            ("clipboard", "get", "-format", "TEXT"),
        ):
            try:
                text = w.tk.call(*cmd)
                if text:
                    return str(text)
            except tk.TclError:
                continue
    return ""


def _first_line(text: str) -> str:
    return text.replace("\r\n", "\n").split("\n", 1)[0].strip()


def _delete_selection(widget: tk.Misc) -> None:
    try:
        cls = widget.winfo_class()
    except tk.TclError:
        return
    if cls == "Text":
        try:
            if widget.tag_ranges(tk.SEL):
                widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
        except tk.TclError:
            pass
        return
    try:
        if widget.selection_present():
            widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
    except tk.TclError:
        pass


def _widget_is_text_disabled(widget: tk.Misc) -> bool:
    try:
        return str(widget.cget("state")) in (tk.DISABLED, "disabled")
    except tk.TclError:
        return False


def _entry_is_editable(widget: tk.Misc) -> bool:
    try:
        state = str(widget.cget("state"))
    except tk.TclError:
        return True
    return state not in (tk.DISABLED, "disabled", "readonly")


def _resolve_paste_target(widget: tk.Misc) -> tuple[tk.Misc, str] | None:
    """
    Найти виджет для вставки. Внутренний Entry у ttk.Combobox не обрабатываем как обычное поле.
    """
    w: tk.Misc | None = widget
    for _ in range(10):
        if w is None:
            return None
        try:
            cls = w.winfo_class()
        except tk.TclError:
            return None
        if cls == "TCombobox":
            return w, "combobox"
        if cls == "Text":
            if not _widget_is_text_disabled(w):
                return w, "text"
            return None
        if cls in ("TEntry", "Entry"):
            try:
                parent = w.master
                if parent is not None and parent.winfo_class() == "TCombobox":
                    return parent, "combobox"
            except tk.TclError:
                pass
            if _entry_is_editable(w):
                return w, "entry"
            return None
        try:
            w = w.master
        except tk.TclError:
            return None
    return None


def _paste_once(widget: tk.Misc, action) -> bool:
    """True — вставка выполнена; False — повтор того же нажатия (уже обработано)."""
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
    try:
        widget.insert(tk.INSERT, text)
    except tk.TclError:
        try:
            widget.insert("insert", text)
        except tk.TclError:
            pass


def _paste_into_combobox(widget: tk.Misc) -> None:
    raw = _clipboard_text(widget)
    if not raw:
        return
    text = _first_line(raw)
    if not text:
        return
    try:
        state = str(widget.cget("state"))
    except tk.TclError:
        state = ""
    if state == "readonly":
        try:
            widget.set(text)
        except tk.TclError:
            pass
        return
    _delete_selection(widget)
    try:
        widget.insert(tk.INSERT, text)
    except tk.TclError:
        try:
            widget.set(((widget.get() or "") + text).strip())
        except tk.TclError:
            pass


def _paste_into_text(widget: tk.Misc) -> None:
    if _widget_is_text_disabled(widget):
        return
    text = _clipboard_text(widget)
    if not text:
        return
    _delete_selection(widget)
    widget.insert(tk.INSERT, text)


def _paste_target(widget: tk.Misc, kind: str) -> None:
    if kind == "combobox":
        _paste_into_combobox(widget)
    elif kind == "text":
        _paste_into_text(widget)
    else:
        _paste_into_entry(widget)


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


def _on_paste(event: tk.Event) -> str | None:
    resolved = _resolve_paste_target(event.widget)
    if resolved is None:
        return None
    target, kind = resolved

    def _do() -> None:
        _paste_target(target, kind)

    if _paste_once(target, _do):
        return "break"
    return None


def _on_copy(event: tk.Event) -> str | None:
    resolved = _resolve_paste_target(event.widget)
    if resolved is None:
        return None
    target, kind = resolved
    if kind == "text":
        _copy_from_text(target)
    else:
        _copy_from_entry(target)
    return "break"


def _on_cut(event: tk.Event) -> str | None:
    resolved = _resolve_paste_target(event.widget)
    if resolved is None:
        return None
    target, kind = resolved
    if kind == "text":
        _cut_from_text(target)
    else:
        _cut_from_entry(target)
    return "break"


def _on_select_all(event: tk.Event) -> str | None:
    resolved = _resolve_paste_target(event.widget)
    if resolved is None:
        return None
    target, kind = resolved
    if kind == "text":
        _select_all_text(target)
    else:
        _select_all_entry(target)
    return "break"


def _popup_context_menu(event: tk.Event) -> str | None:
    resolved = _resolve_paste_target(event.widget)
    if resolved is None:
        return None
    w, kind = resolved
    menu = tk.Menu(w, tearoff=0)

    if kind == "text" and _widget_is_text_disabled(w):
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

    paste_fn = lambda: _paste_once(w, lambda: _paste_target(w, kind))
    if kind == "text":
        copy_fn = lambda: _copy_from_text(w)
        cut_fn = lambda: _cut_from_text(w)
        select_fn = lambda: _select_all_text(w)
    else:
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


def _bind_class_keys(root: tk.Misc, class_name: str) -> None:
    for seq in _paste_sequences_for_class(class_name):
        root.bind_class(class_name, seq, _on_paste)
    for seq, handler in (
        ("<Control-c>", _on_copy),
        ("<Control-C>", _on_copy),
        ("<Control-Insert>", _on_copy),
        ("<Control-x>", _on_cut),
        ("<Control-X>", _on_cut),
        ("<Shift-Delete>", _on_cut),
        ("<Control-a>", _on_select_all),
        ("<Control-A>", _on_select_all),
        ("<Button-3>", _popup_context_menu),
    ):
        root.bind_class(class_name, seq, handler)


def _bind_direct_on_widget(widget: tk.Misc) -> None:
    """Доп. привязка на конкретный виджет (дочерние Entry в ttk, новые Toplevel)."""
    resolved = _resolve_paste_target(widget)
    if resolved is None:
        return
    target, _kind = resolved
    key = id(target)
    if key in _DIRECT_BOUND:
        return
    _DIRECT_BOUND.add(key)

    for seq in _paste_sequences_for_widget(target):
        widget.bind(seq, _on_paste, add="+")
        target.bind(seq, _on_paste, add="+")
    for seq, handler in (
        ("<Control-c>", _on_copy),
        ("<Control-C>", _on_copy),
        ("<Control-x>", _on_cut),
        ("<Control-X>", _on_cut),
        ("<Control-a>", _on_select_all),
        ("<Control-A>", _on_select_all),
        ("<Button-3>", _popup_context_menu),
    ):
        widget.bind(seq, handler, add="+")
        target.bind(seq, handler, add="+")


def _walk_bind_widgets(parent: tk.Misc) -> None:
    try:
        children = parent.winfo_children()
    except tk.TclError:
        return
    for child in children:
        _bind_direct_on_widget(child)
        _walk_bind_widgets(child)


def _register_toplevel_tree(root: tk.Misc) -> None:
    """Обойти корень и все дочерние Toplevel (окна, созданные до/после install)."""
    _walk_bind_widgets(root)
    try:
        for child in root.winfo_children():
            _walk_bind_widgets(child)
    except tk.TclError:
        pass


def bind_editable_clipboard(widget: tk.Misc) -> None:
    """Явная привязка вставки к полю (например, tk.Text во вложенном окне)."""
    _bind_direct_on_widget(widget)


def register_clipboard_window(window: tk.Misc) -> None:
    """Привязать вставку ко всем полям ввода в окне (в т.ч. во вложенных фреймах)."""
    if window is None:
        return
    try:
        top = window.winfo_toplevel()
    except tk.TclError:
        return
    if not _CLIPBOARD_INSTALLED:
        install_clipboard_support(top)
        return
    _walk_bind_widgets(window)
    try:
        window.bind(
            "<Map>",
            lambda e, w=window: _walk_bind_widgets(w),
            add="+",
        )
    except tk.TclError:
        pass


def _bind_global_paste_fallback(root: tk.Misc) -> None:
    """Запасной путь: фокусный виджет, если привязка класса не сработала."""

    def _route(event: tk.Event) -> str | None:
        w = root.focus_get()
        if w is None:
            return None
        resolved = _resolve_paste_target(w)
        if resolved is None:
            return None
        target, kind = resolved

        def _do() -> None:
            _paste_target(target, kind)

        if _paste_once(target, _do):
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
        register_clipboard_window(root)
        return
    _CLIPBOARD_INSTALLED = True

    for class_name in ("TEntry", "Entry", "TCombobox", "Text"):
        _bind_class_keys(root, class_name)

    _bind_global_paste_fallback(root)
    _register_toplevel_tree(root)
