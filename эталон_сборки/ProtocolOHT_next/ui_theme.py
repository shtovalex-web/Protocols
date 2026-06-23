# -*- coding: utf-8 -*-
"""Единая тема интерфейса (палитра и стили как в grafik-pz)."""

from __future__ import annotations

import sys
import tkinter as tk
import tkinter.font as tkfont
from dataclasses import dataclass
from tkinter import ttk

SPACING = 8
SPACING_SM = 4
SPACING_LG = 12
SPACING_XL = 16

_FONT_CANDIDATES = (
    "Segoe UI",
    "DejaVu Sans",
    "Ubuntu",
    "Cantarell",
    "Noto Sans",
    "Liberation Sans",
    "TkDefaultFont",
)


@dataclass(frozen=True)
class Colors:
    """Цвета как в grafik-pz/desktop/ui_style.py."""

    window_bg: str = "#FFFFFF"
    surface: str = "#F3F6FA"
    heading_bg: str = "#D6E4F0"
    heading_fg: str = "#1A237E"
    text: str = "#1A237E"
    text_body: str = "#212121"
    text_muted: str = "#546E7A"
    text_hint: str = "#5C6BC0"
    link: str = "#0563C1"
    border: str = "#90A4AE"
    field_info: str = "#F5FAFF"
    field_date: str = "#FFFBF0"
    tooltip_bg: str = "#FFFFFF"
    tooltip_fg: str = "#1A237E"
    tooltip_border: str = "#90A4AE"
    status: str = "#37474F"
    accent: str = "#1A237E"
    accent_hover: str = "#283593"
    accent_disabled: str = "#9FA8DA"
    search_hit: str = "#FFF3CD"
    success: str = "#2E7D32"
    warning: str = "#F57C00"
    error: str = "#C62828"
    list_select_bg: str = "#1A237E"
    heading_active: str = "#B8D4EC"


@dataclass
class UIState:
    family: str = "TkDefaultFont"
    font_body: tuple[str, int] = ("TkDefaultFont", 9)
    font_small: tuple[str, int] = ("TkDefaultFont", 8)
    font_status: tuple[str, int] = ("TkDefaultFont", 9)
    font_title: tuple[str, int] = ("TkDefaultFont", 10, "bold")
    font_heading: tuple[str, int] = ("TkDefaultFont", 9, "bold")
    font_dialog_title: tuple[str, int] = ("TkDefaultFont", 12, "bold")
    font_mono: tuple[str, int] = ("Consolas", 9)
    font_preview_body: tuple[str, int] = ("TkDefaultFont", 10)


UI = UIState()


def _resolve_font_family(root: tk.Misc | None = None) -> str:
    if sys.platform == "win32":
        try:
            if root is not None:
                tkfont.Font(root=root, family="Segoe UI")
            else:
                tkfont.Font(family="Segoe UI")
            return "Segoe UI"
        except tk.TclError:
            pass
    for name in _FONT_CANDIDATES:
        try:
            if root is not None:
                tkfont.Font(root=root, family=name)
            else:
                tkfont.Font(family=name)
            return name
        except tk.TclError:
            continue
    return "TkDefaultFont"


def _mono_family(root: tk.Misc | None = None) -> str:
    for name in ("Consolas", "DejaVu Sans Mono", "Liberation Mono", "Courier New"):
        try:
            if root is not None:
                tkfont.Font(root=root, family=name)
            else:
                tkfont.Font(family=name)
            return name
        except tk.TclError:
            continue
    return _resolve_font_family(root)


def refresh_fonts(root: tk.Misc | None = None) -> None:
    family = _resolve_font_family(root)
    mono = _mono_family(root)
    UI.family = family
    UI.font_body = (family, 9)
    UI.font_small = (family, 8)
    UI.font_status = (family, 9)
    UI.font_title = (family, 10, "bold")
    UI.font_heading = (family, 9, "bold")
    UI.font_dialog_title = (family, 12, "bold")
    UI.font_mono = (mono, 9)
    UI.font_preview_body = (family, 10)


def pad(*, small: bool = False, large: bool = False) -> dict[str, int]:
    if small:
        p = SPACING_SM
    elif large:
        p = SPACING_LG
    else:
        p = SPACING
    return {"padx": p, "pady": p}


def _configure_clam_theme(style: ttk.Style, root: tk.Misc) -> None:
    f = UI.family
    c = Colors

    style.configure(".", font=UI.font_body, background=c.window_bg)
    style.configure("TFrame", background=c.window_bg)
    style.configure("Toolbar.TFrame", background=c.heading_bg)
    style.configure("TLabel", font=UI.font_body, foreground=c.text_body, background=c.window_bg)
    style.configure("Toolbar.TLabel", background=c.heading_bg, foreground=c.heading_fg)
    style.configure("Hint.TLabel", font=UI.font_small, foreground=c.text_hint, background=c.window_bg)
    style.configure("Muted.TLabel", font=UI.font_body, foreground=c.text_muted, background=c.window_bg)
    style.configure("Status.TLabel", font=UI.font_status, foreground=c.status, background=c.window_bg)
    style.configure("Title.TLabel", font=UI.font_title, foreground=c.heading_fg, background=c.window_bg)

    style.configure("TButton", font=UI.font_body, padding=(10, 6))
    style.configure("Small.TButton", font=UI.font_small, padding=(8, 4))
    style.configure(
        "Accent.TButton",
        font=UI.font_heading,
        padding=(12, 7),
        background=c.accent,
        foreground="#FFFFFF",
        borderwidth=0,
        focusthickness=2,
        focuscolor=c.accent,
    )
    style.map(
        "Accent.TButton",
        background=[
            ("active", c.accent_hover),
            ("pressed", c.accent_hover),
            ("disabled", c.accent_disabled),
        ],
        foreground=[("disabled", "#ECEFF1")],
    )
    style.map("TButton", background=[("active", c.heading_active)])

    style.configure("TLabelframe", padding=SPACING, background=c.window_bg, bordercolor=c.border)
    style.configure(
        "TLabelframe.Label",
        font=UI.font_heading,
        foreground=c.heading_fg,
        background=c.heading_bg,
    )
    style.configure("Card.TLabelframe", padding=SPACING_LG, background=c.window_bg)
    style.configure(
        "Card.TLabelframe.Label",
        font=UI.font_title,
        foreground=c.heading_fg,
        background=c.heading_bg,
    )

    style.configure("TEntry", font=UI.font_body, fieldbackground=c.field_info, padding=4)
    style.configure("Field.TEntry", font=UI.font_body, fieldbackground=c.field_info, padding=4)
    style.configure("FieldDate.TEntry", font=UI.font_body, fieldbackground=c.field_date, padding=4)
    style.configure("TCombobox", font=UI.font_body, fieldbackground=c.field_info, padding=4)
    style.configure("Toolbutton", padding=(6, 4))
    style.configure("TSeparator", background=c.border)

    style.configure("Treeview", rowheight=26, font=UI.font_body, background=c.window_bg)
    style.configure(
        "Treeview.Heading",
        font=UI.font_heading,
        background=c.heading_bg,
        foreground=c.heading_fg,
        relief="flat",
        padding=(6, 4),
    )
    style.map("Treeview.Heading", background=[("active", c.heading_active)])

    try:
        root.configure(background=c.window_bg)
    except tk.TclError:
        pass


def apply_theme(root: tk.Misc) -> ttk.Style:
    """Настроить ttk и шрифты (clam + палитра grafik-pz)."""
    refresh_fonts(root)
    style = ttk.Style(root)
    try:
        if "clam" in style.theme_names():
            style.theme_use("clam")
        elif sys.platform == "win32":
            style.theme_use("vista")
        elif sys.platform == "darwin":
            style.theme_use("aqua")
    except tk.TclError:
        pass

    if style.theme_use() == "clam":
        _configure_clam_theme(style, root)
    else:
        f = UI.family
        style.configure(".", font=UI.font_body)
        style.configure("Hint.TLabel", font=UI.font_small, foreground=Colors.text_hint)
        style.configure("Accent.TButton", font=(f, 10, "bold"), padding=(14, 8))

    return style


def apply_startup_geometry(
    win: tk.Misc,
    *,
    min_width: int = 900,
    min_height: int = 560,
) -> None:
    """Подогнать и отцентрировать главное окно при старте (без полноэкранного режима)."""
    win.minsize(min_width, min_height)
    _center_window(
        win,
        min_width=min(min_width + 80, win.winfo_screenwidth() - 40),
        min_height=min_height,
    )


def _center_window(win: tk.Misc, *, min_width: int, min_height: int) -> None:
    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    width = min(max(min_width, win.winfo_reqwidth() + 16), sw - 32)
    height = min(max(min_height, win.winfo_reqheight() + 16), sh - 48)
    x = max(0, (sw - width) // 2)
    y = max(0, (sh - height) // 2)
    win.geometry(f"{width}x{height}+{x}+{y}")


def configure_readonly_text(widget: tk.Text, *, mono: bool = False) -> None:
    widget.configure(
        font=UI.font_mono if mono else UI.font_body,
        padx=SPACING,
        pady=SPACING,
        relief=tk.FLAT,
        borderwidth=1,
        highlightthickness=1,
        highlightbackground=Colors.border,
        highlightcolor=Colors.accent,
        background=Colors.window_bg,
        foreground=Colors.text_body,
        insertbackground=Colors.text_body,
    )


def configure_listbox(widget: tk.Listbox, *, mono: bool = False) -> None:
    widget.configure(
        font=UI.font_mono if mono else UI.font_body,
        relief=tk.FLAT,
        borderwidth=1,
        highlightthickness=1,
        highlightbackground=Colors.border,
        highlightcolor=Colors.accent,
        background=Colors.window_bg,
        foreground=Colors.text_body,
        selectbackground=Colors.list_select_bg,
        selectforeground="#FFFFFF",
        activestyle="none",
    )
