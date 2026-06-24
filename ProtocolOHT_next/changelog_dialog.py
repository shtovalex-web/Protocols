# -*- coding: utf-8 -*-
"""Диалог «Что нового» после обновления."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from protocol_app_info import APP_FULL_NAME


def show_changelog_dialog(version: str, changes: list[str]) -> None:
    root = tk.Tk()
    root.withdraw()

    dialog = tk.Toplevel(root)
    dialog.title("Что нового")
    dialog.resizable(False, False)
    dialog.transient(root)
    dialog.grab_set()

    frame = ttk.Frame(dialog, padding=16)
    frame.pack(fill=tk.BOTH, expand=True)

    ttk.Label(frame, text=APP_FULL_NAME, font=("", 11, "bold")).pack(anchor="w")
    ttk.Label(frame, text=f"Обновлено до версии {version}").pack(anchor="w", pady=(4, 12))

    if changes:
        body = "\n".join(f"• {item}" for item in changes)
    else:
        body = "Список изменений недоступен."
    ttk.Label(frame, text=body, justify=tk.LEFT, wraplength=420).pack(anchor="w", pady=(0, 12))

    ttk.Button(frame, text="Продолжить", command=dialog.destroy).pack(anchor="e")

    dialog.update_idletasks()
    dialog.geometry(
        f"+{(root.winfo_screenwidth() - dialog.winfo_width()) // 2}+"
        f"{(root.winfo_screenheight() - dialog.winfo_height()) // 2}"
    )
    dialog.wait_window()
    root.destroy()
