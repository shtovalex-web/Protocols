# -*- coding: utf-8 -*-
"""Точка входа: формирование протоколов проверки знаний по охране труда (tkinter).

Усовершенствованный модульный код хранится в папке ProtocolOHT_next/ (protocol_ui, protocol_docx,
protocol_journal и др.). Корневой main.py только подключает эту папку к sys.path и запускает приложение.
"""

from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent
_NEXT = _ROOT / "ProtocolOHT_next"
# Удаляем дубликаты порядка, затем: сначала ProtocolOHT_next, затем корень (commission_admin, employees_io…).
for _p in (str(_ROOT), str(_NEXT)):
    while _p in sys.path:
        sys.path.remove(_p)
sys.path.insert(0, str(_NEXT))
sys.path.insert(1, str(_ROOT))

from protocol_db import init_db
from protocol_errors import append_error_journal, install_app_error_logging
from protocol_paths import ensure_frozen_default_workbooks, migrate_legacy_from_data_subfolder_to_exe
from protocol_ui import ProtocolApp


def main() -> None:
    from startup_update import prepare_startup_updates

    if not prepare_startup_updates(sys.argv):
        return

    migrate_legacy_from_data_subfolder_to_exe()
    ensure_frozen_default_workbooks()
    install_app_error_logging()
    journal_removed = init_db()
    try:
        app = ProtocolApp(journal_duplicates_removed=journal_removed)
    except Exception as e:
        import traceback

        tb_text = traceback.format_exc()
        try:
            append_error_journal(
                "Критическая ошибка при запуске (ProtocolApp)",
                str(e),
                traceback_text=tb_text,
            )
        except Exception:
            pass
        try:
            import tkinter as tk
            from tkinter import messagebox

            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(
                "Не удалось запустить программу",
                f"{type(e).__name__}: {e}\n\n"
                "Подробности — в файле protocol_errors_journal.txt рядом с программой "
                "(или в папке пользовательских данных при сборке .exe).",
                parent=root,
            )
            root.destroy()
        except Exception:
            traceback.print_exc()
        sys.exit(1)
    app.mainloop()


if __name__ == "__main__":
    main()
