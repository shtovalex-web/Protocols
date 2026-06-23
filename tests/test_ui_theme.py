# -*- coding: utf-8 -*-
"""Smoke-тесты единой темы интерфейса."""

from __future__ import annotations

import tkinter as tk
import unittest
from tkinter import ttk

from _bootstrap import setup_main_project_paths

setup_main_project_paths()

from ui_theme import (
    CORNER_RADIUS,
    FIELD_DATE_STYLE,
    FIELD_STYLE,
    SETTING_UI_COLOR_SCHEME,
    Colors,
    UI,
    apply_color_scheme,
    apply_theme,
    apply_theme_to_window,
    color_scheme_choices,
    configure_listbox,
    current_color_scheme_id,
    load_color_scheme_from_settings,
    pad,
    refresh_fonts,
    save_color_scheme_to_settings,
    set_color_scheme,
)


class TestUITheme(unittest.TestCase):
    def test_refresh_fonts_sets_body_tuple(self) -> None:
        root = tk.Tk()
        root.withdraw()
        try:
            refresh_fonts(root)
            self.assertIsInstance(UI.font_body, tuple)
            self.assertEqual(len(UI.font_body), 2)
            self.assertEqual(UI.font_body[1], 9)
        finally:
            root.destroy()

    def test_apply_theme_configures_ttk_styles(self) -> None:
        root = tk.Tk()
        root.withdraw()
        try:
            style = apply_theme(root)
            self.assertEqual(style.lookup("Hint.TLabel", "foreground"), Colors.text_hint)
            self.assertEqual(style.theme_use(), "clam")
            self.assertEqual(style.lookup("Accent.TButton", "background"), Colors.accent)
            self.assertEqual(style.lookup(FIELD_STYLE, "fieldbackground"), Colors.field_info)
            self.assertEqual(style.lookup(FIELD_DATE_STYLE, "fieldbackground"), Colors.field_date)
            lbl = ttk.Label(root, text="x", style="Hint.TLabel")
            self.assertEqual(lbl.cget("style"), "Hint.TLabel")
        finally:
            root.destroy()

    def test_apply_theme_to_window_on_toplevel(self) -> None:
        root = tk.Tk()
        root.withdraw()
        try:
            apply_theme(root)
            win = tk.Toplevel(root)
            apply_theme_to_window(win)
            self.assertEqual(win.cget("background"), Colors.window_bg)
            ent = ttk.Entry(win, style=FIELD_STYLE)
            self.assertEqual(ent.cget("style"), FIELD_STYLE)
        finally:
            try:
                win.destroy()
            except Exception:
                pass
            root.destroy()

    def test_corner_radius_positive(self) -> None:
        self.assertGreaterEqual(CORNER_RADIUS, 4)

    def test_color_scheme_switch_changes_accent(self) -> None:
        set_color_scheme("grafik")
        blue = Colors.accent
        set_color_scheme("forest")
        self.assertNotEqual(Colors.accent, blue)
        self.assertEqual(current_color_scheme_id(), "forest")

    def test_color_scheme_choices_not_empty(self) -> None:
        self.assertGreaterEqual(len(color_scheme_choices()), 2)

    def test_tbutton_has_visible_background(self) -> None:
        root = tk.Tk()
        root.withdraw()
        try:
            set_color_scheme("grafik")
            style = apply_theme(root)
            bg = style.lookup("TButton", "background")
            self.assertTrue(bg)
            self.assertNotEqual(bg, Colors.window_bg)
            border = style.lookup("TButton", "bordercolor") or Colors.button_border
            self.assertNotEqual(border, Colors.window_bg)
            self.assertEqual(style.lookup("TButton", "borderwidth"), 2)
        finally:
            root.destroy()

    def test_apply_color_scheme_persists_without_db(self) -> None:
        root = tk.Tk()
        root.withdraw()
        try:
            apply_color_scheme("warm", root, persist=False)
            self.assertEqual(current_color_scheme_id(), "warm")
            self.assertEqual(Colors.accent, "#E65100")
        finally:
            root.destroy()

    def test_color_scheme_persists_in_app_settings(self) -> None:
        import tempfile
        from pathlib import Path

        from _bootstrap import close_tracked_sqlite_connections, enable_sqlite_test_tracking
        from protocol_db import init_protocols_db_file

        enable_sqlite_test_tracking()
        with tempfile.TemporaryDirectory() as tmp:
            db_path = Path(tmp) / "protocols.db"
            init_protocols_db_file(db_path)
            import commission_admin as ca

            orig = ca.database_path
            ca.database_path = lambda: db_path  # type: ignore[assignment]
            try:
                set_color_scheme("grafik")
                save_color_scheme_to_settings("forest")
                set_color_scheme("grafik")
                load_color_scheme_from_settings()
                self.assertEqual(current_color_scheme_id(), "forest")
                self.assertEqual(
                    ca._app_setting_get(SETTING_UI_COLOR_SCHEME, ""),
                    "forest",
                )
            finally:
                ca.database_path = orig  # type: ignore[assignment]
                close_tracked_sqlite_connections()

    def test_pad_returns_symmetric_spacing(self) -> None:
        small = pad(small=True)
        normal = pad()
        self.assertLess(small["padx"], normal["padx"])
        self.assertEqual(normal["padx"], normal["pady"])

    def test_configure_listbox_applies_font(self) -> None:
        root = tk.Tk()
        root.withdraw()
        try:
            refresh_fonts(root)
            lb = tk.Listbox(root)
            configure_listbox(lb, mono=True)
            font_val = lb.cget("font")
            if isinstance(font_val, str):
                self.assertIn(UI.font_mono[0], font_val)
            else:
                self.assertEqual(font_val, UI.font_mono)
        finally:
            root.destroy()


if __name__ == "__main__":
    unittest.main()
