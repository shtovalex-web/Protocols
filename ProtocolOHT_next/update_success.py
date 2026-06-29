# -*- coding: utf-8 -*-
"""Уведомление об успешном обновлении и выход без автоперезапуска."""

from __future__ import annotations

from pathlib import Path

from tkinter import messagebox

from update_installer import exit_after_update


def notify_update_success_and_exit(
    *,
    version: str,
    exe_path: Path,
    parent=None,
    bundle_staged: bool = False,
) -> None:
    exe_name = exe_path.name
    if bundle_staged:
        restart_lines = (
            f"Закройте программу и снова запустите {exe_name}.\n"
            "При следующем запуске файлы программы будут заменены."
        )
    else:
        restart_lines = f"Закройте программу и снова запустите {exe_name}."
    messagebox.showinfo(
        "Обновление",
        f"Обновление до версии {version} прошло успешно.\n\n{restart_lines}",
        parent=parent,
    )
    if parent is not None:
        try:
            top = parent.winfo_toplevel()
            top.quit()
        except Exception:
            pass
    exit_after_update()
