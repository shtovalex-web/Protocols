# -*- coding: utf-8 -*-
"""Единая тема интерфейса: палитры, стили полей и кнопок."""

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

CORNER_RADIUS = 8
CORNER_HIGHLIGHT = 2

FIELD_STYLE = "Field.TEntry"
FIELD_DATE_STYLE = "FieldDate.TEntry"
FIELD_COMBO_STYLE = "Field.TCombobox"

SETTING_UI_COLOR_SCHEME = "ui_color_scheme"
DEFAULT_SCHEME_ID = "grafik"

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
class ColorScheme:
    """Палитра интерфейса."""

    window_bg: str
    surface: str
    heading_bg: str
    heading_fg: str
    text: str
    text_body: str
    text_muted: str
    text_hint: str
    link: str
    border: str
    field_info: str
    field_date: str
    tooltip_bg: str
    tooltip_fg: str
    tooltip_border: str
    status: str
    accent: str
    accent_hover: str
    accent_disabled: str
    button_bg: str
    button_border: str
    button_active: str
    search_hit: str
    success: str
    warning: str
    error: str
    list_select_bg: str
    heading_active: str


def _scheme(
    *,
    window_bg: str = "#FFFFFF",
    surface: str = "#F3F6FA",
    heading_bg: str = "#D6E4F0",
    heading_fg: str = "#1A237E",
    text: str = "#1A237E",
    text_body: str = "#212121",
    text_muted: str = "#546E7A",
    text_hint: str = "#5C6BC0",
    link: str = "#0563C1",
    border: str = "#90A4AE",
    field_info: str = "#F5FAFF",
    field_date: str = "#FFFBF0",
    tooltip_bg: str = "#FFFFFF",
    tooltip_fg: str = "#1A237E",
    tooltip_border: str = "#90A4AE",
    status: str = "#37474F",
    accent: str = "#1A237E",
    accent_hover: str = "#283593",
    accent_disabled: str = "#9FA8DA",
    button_bg: str = "#A8C8E8",
    button_border: str = "#37474F",
    button_active: str = "#7EB3DD",
    search_hit: str = "#FFF3CD",
    success: str = "#2E7D32",
    warning: str = "#F57C00",
    error: str = "#C62828",
    list_select_bg: str = "#1A237E",
    heading_active: str = "#B8D4EC",
) -> ColorScheme:
    return ColorScheme(
        window_bg=window_bg,
        surface=surface,
        heading_bg=heading_bg,
        heading_fg=heading_fg,
        text=text,
        text_body=text_body,
        text_muted=text_muted,
        text_hint=text_hint,
        link=link,
        border=border,
        field_info=field_info,
        field_date=field_date,
        tooltip_bg=tooltip_bg,
        tooltip_fg=tooltip_fg,
        tooltip_border=tooltip_border,
        status=status,
        accent=accent,
        accent_hover=accent_hover,
        accent_disabled=accent_disabled,
        button_bg=button_bg,
        button_border=button_border,
        button_active=button_active,
        search_hit=search_hit,
        success=success,
        warning=warning,
        error=error,
        list_select_bg=list_select_bg,
        heading_active=heading_active,
    )


COLOR_SCHEMES: dict[str, ColorScheme] = {
    "grafik": _scheme(),
    "forest": _scheme(
        heading_bg="#C8E6C9",
        heading_fg="#1B5E20",
        text="#1B5E20",
        text_hint="#388E3C",
        accent="#2E7D32",
        accent_hover="#1B5E20",
        accent_disabled="#A5D6A7",
        list_select_bg="#2E7D32",
        field_info="#F1F8E9",
        field_date="#FFFDE7",
        surface="#E8F5E9",
        button_bg="#8BC34A",
        button_border="#1B5E20",
        button_active="#689F38",
        heading_active="#A5D6A7",
        link="#1B5E20",
        tooltip_fg="#1B5E20",
    ),
    "slate": _scheme(
        window_bg="#ECEFF1",
        surface="#E0E4E8",
        heading_bg="#CFD8DC",
        heading_fg="#263238",
        text="#263238",
        text_hint="#546E7A",
        accent="#455A64",
        accent_hover="#37474F",
        accent_disabled="#B0BEC5",
        list_select_bg="#455A64",
        field_info="#F5F7F8",
        field_date="#FFF8E1",
        button_bg="#90A4AE",
        button_border="#263238",
        button_active="#607D8B",
        heading_active="#B0BEC5",
        link="#1565C0",
        tooltip_fg="#263238",
    ),
    "warm": _scheme(
        surface="#FFF3E0",
        heading_bg="#FFE0B2",
        heading_fg="#BF360C",
        text="#BF360C",
        text_hint="#E65100",
        accent="#E65100",
        accent_hover="#BF360C",
        accent_disabled="#FFCC80",
        list_select_bg="#E65100",
        field_info="#FFF8F0",
        field_date="#FFFDE7",
        button_bg="#FFB74D",
        button_border="#BF360C",
        button_active="#FFA726",
        heading_active="#FFCC80",
        link="#E65100",
        tooltip_fg="#BF360C",
    ),
}

SCHEME_LABELS: dict[str, str] = {
    "grafik": "Классическая (синяя)",
    "forest": "Зелёная",
    "slate": "Серая",
    "warm": "Тёплая",
}

_active_scheme_id: str = DEFAULT_SCHEME_ID
_active: ColorScheme = COLOR_SCHEMES[DEFAULT_SCHEME_ID]


class _ColorsProxy:
    """Текущая палитра (меняется при смене цветовой схемы)."""

    def __getattr__(self, name: str) -> str:
        return getattr(_active, name)


Colors = _ColorsProxy()


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


def color_scheme_choices() -> list[tuple[str, str]]:
    return [(sid, SCHEME_LABELS.get(sid, sid)) for sid in COLOR_SCHEMES]


def current_color_scheme_id() -> str:
    return _active_scheme_id


def set_color_scheme(scheme_id: str) -> ColorScheme:
    global _active_scheme_id, _active
    sid = (scheme_id or "").strip() or DEFAULT_SCHEME_ID
    if sid not in COLOR_SCHEMES:
        sid = DEFAULT_SCHEME_ID
    _active_scheme_id = sid
    _active = COLOR_SCHEMES[sid]
    return _active


def load_color_scheme_from_settings() -> str:
    try:
        from commission_admin import _app_setting_get

        raw = _app_setting_get(SETTING_UI_COLOR_SCHEME, DEFAULT_SCHEME_ID).strip()
    except Exception:
        raw = DEFAULT_SCHEME_ID
    set_color_scheme(raw)
    return _active_scheme_id


def save_color_scheme_to_settings(scheme_id: str) -> None:
    from commission_admin import _app_setting_set

    set_color_scheme(scheme_id)
    _app_setting_set(SETTING_UI_COLOR_SCHEME, _active_scheme_id)


def apply_color_scheme(scheme_id: str, root: tk.Misc, *, persist: bool = True) -> ttk.Style:
    set_color_scheme(scheme_id)
    if persist:
        save_color_scheme_to_settings(_active_scheme_id)
    return apply_theme(root)


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


def _rounded_tk_chrome(widget: tk.Widget, *, bg: str | None = None) -> None:
    c = Colors
    fill = bg if bg is not None else c.window_bg
    try:
        widget.configure(
            relief=tk.FLAT,
            borderwidth=0,
            highlightthickness=CORNER_HIGHLIGHT,
            highlightbackground=c.border,
            highlightcolor=c.accent,
            background=fill,
        )
    except tk.TclError:
        pass


def _configure_buttons(style: ttk.Style) -> None:
    c = Colors
    for name, padding, font in (
        ("TButton", (12, 7), UI.font_body),
        ("Small.TButton", (8, 5), UI.font_small),
    ):
        style.configure(
            name,
            font=font,
            padding=padding,
            background=c.button_bg,
            foreground=c.text,
            borderwidth=2,
            relief="raised",
            bordercolor=c.button_border,
            lightcolor=c.button_border,
            darkcolor=c.button_border,
            focusthickness=2,
            focuscolor=c.accent,
        )
        style.map(
            name,
            background=[
                ("active", c.button_active),
                ("pressed", c.heading_active),
                ("disabled", c.surface),
            ],
            foreground=[("disabled", c.text_muted)],
            bordercolor=[
                ("focus", c.accent),
                ("active", c.accent),
                ("!focus", c.button_border),
            ],
            lightcolor=[("!disabled", c.button_border)],
            darkcolor=[("!disabled", c.button_border)],
            relief=[("pressed", "sunken"), ("!pressed", "raised")],
        )
    style.configure(
        "Accent.TButton",
        font=UI.font_heading,
        padding=(14, 8),
        background=c.accent,
        foreground="#FFFFFF",
        borderwidth=2,
        relief="raised",
        bordercolor=c.accent_hover,
        lightcolor=c.accent_hover,
        darkcolor=c.accent_hover,
        focusthickness=2,
        focuscolor="#FFFFFF",
    )
    style.map(
        "Accent.TButton",
        background=[
            ("active", c.accent_hover),
            ("pressed", c.accent_hover),
            ("disabled", c.accent_disabled),
        ],
        foreground=[("disabled", "#ECEFF1")],
        bordercolor=[
            ("active", c.accent_hover),
            ("!focus", c.accent_hover),
        ],
        lightcolor=[("!disabled", c.accent_hover)],
        darkcolor=[("!disabled", c.accent_hover)],
        relief=[("pressed", "sunken"), ("!pressed", "raised")],
    )
    style.configure(
        "Actions.TFrame",
        background=c.surface,
    )


def _configure_soft_corners(style: ttk.Style) -> None:
    c = Colors
    pad_entry = (8, 6)
    for name in ("TEntry", FIELD_STYLE, FIELD_DATE_STYLE):
        style.configure(name, padding=pad_entry, borderwidth=1, relief="flat")
    style.configure(FIELD_COMBO_STYLE, padding=pad_entry, fieldbackground=c.field_info)
    style.map(
        "TEntry",
        bordercolor=[("focus", c.accent), ("!focus", c.border)],
        lightcolor=[("focus", c.accent), ("!focus", c.border)],
        darkcolor=[("focus", c.accent), ("!focus", c.border)],
    )
    style.map(
        FIELD_STYLE,
        bordercolor=[("focus", c.accent), ("!focus", c.border)],
        fieldbackground=[("readonly", c.field_info), ("!disabled", c.field_info)],
    )
    style.map(
        FIELD_DATE_STYLE,
        bordercolor=[("focus", c.accent), ("!focus", c.border)],
        fieldbackground=[("readonly", c.field_date), ("!disabled", c.field_date)],
    )
    style.configure("TLabelframe", borderwidth=1, relief="flat", bordercolor=c.border)
    style.configure(
        "Card.TLabelframe",
        borderwidth=1,
        relief="flat",
        bordercolor=c.border,
        padding=SPACING_LG,
        background=c.window_bg,
    )
    style.configure("TNotebook", tabmargins=[CORNER_RADIUS // 2, 4, CORNER_RADIUS // 2, 0])
    style.configure(
        "TNotebook.Tab",
        padding=(14, 6),
        background=c.button_bg,
    )
    style.map(
        "TNotebook.Tab",
        background=[("selected", c.heading_bg), ("active", c.button_active)],
    )


def _configure_clam_theme(style: ttk.Style, root: tk.Misc) -> None:
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

    _configure_buttons(style)

    style.configure("TLabelframe", padding=SPACING, background=c.window_bg, bordercolor=c.border)
    style.configure(
        "TLabelframe.Label",
        font=UI.font_heading,
        foreground=c.heading_fg,
        background=c.heading_bg,
    )
    style.configure(
        "Card.TLabelframe.Label",
        font=UI.font_title,
        foreground=c.heading_fg,
        background=c.heading_bg,
    )

    style.configure("TEntry", font=UI.font_body, fieldbackground=c.field_info, padding=4)
    style.configure(FIELD_STYLE, font=UI.font_body, fieldbackground=c.field_info, padding=4)
    style.configure(FIELD_DATE_STYLE, font=UI.font_body, fieldbackground=c.field_date, padding=4)
    style.configure("TCombobox", font=UI.font_body, fieldbackground=c.field_info, padding=4)
    style.configure(FIELD_COMBO_STYLE, font=UI.font_body, fieldbackground=c.field_info, padding=4)
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

    _configure_soft_corners(style)

    try:
        root.configure(background=c.window_bg)
    except tk.TclError:
        pass


def apply_theme(root: tk.Misc) -> ttk.Style:
    """Настроить ttk и шрифты (clam + активная палитра)."""
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
        style.configure(FIELD_STYLE, fieldbackground=Colors.field_info)
        style.configure(FIELD_DATE_STYLE, fieldbackground=Colors.field_date)
        style.configure(FIELD_COMBO_STYLE, fieldbackground=Colors.field_info)
        _configure_buttons(style)

    return style


def apply_theme_to_window(win: tk.Misc) -> ttk.Style:
    """Применить тему к Toplevel/диалогу (ttk + фон окна)."""
    top = win.winfo_toplevel()
    style = apply_theme(top)
    try:
        win.configure(background=Colors.window_bg)
    except tk.TclError:
        pass
    return style


def window_appears_maximized_or_fullscreen(
    win: tk.Misc,
    *,
    margin_w: int = 64,
    margin_h: int = 96,
) -> bool:
    """Окно развёрнуто WM или занимает почти весь экран (не трогать geometry)."""
    try:
        if str(win.wm_state()).lower() == "zoomed":
            return True
    except tk.TclError:
        pass
    try:
        if bool(win.attributes("-zoomed")):
            return True
    except tk.TclError:
        pass
    try:
        win.update_idletasks()
        sw = int(win.winfo_screenwidth())
        sh = int(win.winfo_screenheight())
        ww = int(win.winfo_width())
        wh = int(win.winfo_height())
    except tk.TclError:
        return False
    if ww < 200 or wh < 200:
        return False
    return ww >= sw - margin_w and wh >= sh - margin_h


def apply_startup_geometry(
    win: tk.Misc,
    *,
    min_width: int = 900,
    min_height: int = 560,
) -> None:
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
        insertbackground=Colors.text_body,
        foreground=Colors.text_body,
    )
    _rounded_tk_chrome(widget, bg=Colors.window_bg)


def configure_editable_text(widget: tk.Text, *, mono: bool = False) -> None:
    widget.configure(
        font=UI.font_mono if mono else UI.font_body,
        padx=SPACING,
        pady=SPACING,
        insertbackground=Colors.text_body,
        foreground=Colors.text_body,
    )
    _rounded_tk_chrome(widget, bg=Colors.field_info)


def configure_listbox(widget: tk.Listbox, *, mono: bool = False) -> None:
    widget.configure(
        font=UI.font_mono if mono else UI.font_body,
        foreground=Colors.text_body,
        selectbackground=Colors.list_select_bg,
        selectforeground="#FFFFFF",
        activestyle="none",
    )
    _rounded_tk_chrome(widget, bg=Colors.window_bg)


def configure_canvas(widget: tk.Canvas, *, bg: str | None = None) -> None:
    _rounded_tk_chrome(widget, bg=bg or Colors.window_bg)
