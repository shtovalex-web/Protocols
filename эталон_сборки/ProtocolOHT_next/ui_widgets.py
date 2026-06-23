# -*- coding: utf-8 -*-
"""Переиспользуемые UI-компоненты поверх ui_theme."""

from __future__ import annotations

from typing import Any, Callable

import tkinter as tk
from tkinter import ttk

from ui_theme import Colors, FIELD_STYLE, SPACING, SPACING_SM, UI, pad


class WidgetTooltip:
    """Всплывающая подсказка при наведении."""

    def __init__(
        self,
        widget: tk.Misc,
        text: str,
        *,
        delay_ms: int = 450,
        wraplength: int = 420,
    ) -> None:
        self._widget = widget
        self._text = (text or "").strip()
        self._delay_ms = delay_ms
        self._wraplength = wraplength
        self._after_id: str | None = None
        self._tip: tk.Toplevel | None = None
        if not self._text:
            return
        widget.bind("<Enter>", self._on_enter, add=True)
        widget.bind("<Leave>", self._on_leave, add=True)
        widget.bind("<ButtonPress>", self._on_leave, add=True)
        widget.bind("<Destroy>", self._on_destroy, add=True)

    def _on_destroy(self, _event: object | None = None) -> None:
        self._hide()

    def _cancel_scheduled(self) -> None:
        if self._after_id is not None:
            try:
                self._widget.after_cancel(self._after_id)
            except (tk.TclError, ValueError):
                pass
            self._after_id = None

    def _hide(self, _event: object | None = None) -> None:
        self._cancel_scheduled()
        if self._tip is not None:
            try:
                self._tip.destroy()
            except tk.TclError:
                pass
            self._tip = None

    def _on_enter(self, _event: object | None = None) -> None:
        self._cancel_scheduled()
        self._after_id = self._widget.after(self._delay_ms, self._show_tip)

    def _on_leave(self, _event: object | None = None) -> None:
        self._hide()

    def _show_tip(self) -> None:
        self._after_id = None
        if not self._text:
            return
        try:
            if not self._widget.winfo_exists():
                return
        except tk.TclError:
            return
        root = self._widget.winfo_toplevel()
        self._tip = tw = tk.Toplevel(root)
        tw.wm_overrideredirect(True)
        try:
            tw.attributes("-topmost", True)
        except tk.TclError:
            pass
        frame = tk.Frame(
            tw,
            background=Colors.tooltip_bg,
            highlightbackground=Colors.tooltip_border,
            highlightthickness=1,
            borderwidth=0,
        )
        frame.pack()
        lbl = tk.Label(
            frame,
            text=self._text,
            justify=tk.LEFT,
            wraplength=self._wraplength,
            foreground=Colors.tooltip_fg,
            background=Colors.tooltip_bg,
            font=UI.font_status,
            padx=SPACING,
            pady=SPACING_SM,
        )
        lbl.pack()
        tw.update_idletasks()
        x = int(self._widget.winfo_rootx() + self._widget.winfo_width() // 2 - tw.winfo_width() // 2)
        y = int(self._widget.winfo_rooty() + self._widget.winfo_height() + 6)
        sw = tw.winfo_screenwidth()
        sh = tw.winfo_screenheight()
        tw_w = tw.winfo_width()
        tw_h = tw.winfo_height()
        if x + tw_w > sw - SPACING:
            x = max(SPACING, sw - tw_w - SPACING)
        if y + tw_h > sh - SPACING:
            y = max(SPACING, int(self._widget.winfo_rooty() - tw_h - 6))
        tw.wm_geometry(f"+{x}+{y}")


def attach_tooltip(widget: tk.Misc, text: str, **kw: Any) -> WidgetTooltip | None:
    if not (text or "").strip():
        return None
    return WidgetTooltip(widget, text, **kw)


def build_dialog_button_row(
    parent: tk.Misc,
    *,
    on_close: Callable[[], None],
    extra_before: tuple[ttk.Button, ...] = (),
) -> ttk.Frame:
    row = ttk.Frame(parent, padding=(0, SPACING, 0, 0))
    row.pack(fill=tk.X)
    for btn in extra_before:
        btn.pack(side=tk.LEFT, padx=(0, SPACING_SM))
    ttk.Button(row, text="Закрыть", command=on_close).pack(side=tk.RIGHT)
    return row


def build_search_toolbar(
    parent: tk.Misc,
    *,
    on_highlight: Callable[[], None],
    on_next: Callable[[], None],
) -> tuple[ttk.Frame, tk.StringVar, ttk.Label, ttk.Entry]:
    bar = ttk.Frame(parent, padding=pad())
    bar.pack(fill=tk.X)
    ttk.Label(bar, text="Поиск (от 2 симв.):").pack(side=tk.LEFT, padx=(0, SPACING_SM))
    var_q = tk.StringVar()
    ent = ttk.Entry(bar, textvariable=var_q, width=32, style=FIELD_STYLE)
    ent.pack(side=tk.LEFT, padx=(0, SPACING_SM))
    ttk.Button(bar, text="Подсветить", command=on_highlight, style="Small.TButton").pack(
        side=tk.LEFT, padx=(0, SPACING_SM)
    )
    ttk.Button(bar, text="Далее", command=on_next, style="Small.TButton").pack(side=tk.LEFT)
    status = ttk.Label(bar, text="", style="Hint.TLabel")
    status.pack(side=tk.LEFT, padx=(SPACING, 0))
    return bar, var_q, status, ent
