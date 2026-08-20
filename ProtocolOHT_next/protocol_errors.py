# -*- coding: utf-8 -*-
"""Текстовый журнал ошибок и перехват messagebox / необработанных исключений."""

from __future__ import annotations

import re
import sys
import threading
import traceback
from datetime import datetime

from app_paths import application_error_log_path
from tkinter import messagebox


_error_journal_lock = threading.Lock()

# СНИЛС в формате XXX-XXX-XXX XX / XXX-XXX-XXX-XX; не трогаем короткие числа.
_SNILS_RE = re.compile(
    r"\b\d{3}[-\s]?\d{3}[-\s]?\d{3}[-\s]?\d{2}\b"
)
# Типичные подписи с ФИО в сообщениях UI.
_FIO_LABEL_RE = re.compile(
    r"(?im)^(\s*(?:сотрудник|фио|работник)\s*[:：]\s*).+$"
)


def sanitize_error_journal_text(text: str) -> str:
    """Убрать из текста журнала СНИЛС и строки с подписью ФИО (152-ФЗ)."""
    s = text or ""
    s = _SNILS_RE.sub("[СНИЛС]", s)
    s = _FIO_LABEL_RE.sub(r"\1[обезличено]", s)
    return s


def append_error_journal(
    title: str,
    message: str,
    *,
    traceback_text: str | None = None,
) -> None:
    """Дозапись в текстовый журнал ошибок (см. application_error_log_path)."""
    try:
        path = application_error_log_path()
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        parts = [
            "=" * 72,
            f"[{stamp}] {sanitize_error_journal_text(title)}",
            sanitize_error_journal_text(message or "").rstrip(),
        ]
        tt = (traceback_text or "").strip()
        if tt:
            parts.append(sanitize_error_journal_text(tt))
        parts.append("")
        block = "\n".join(parts) + "\n"
        with _error_journal_lock:
            path.parent.mkdir(parents=True, exist_ok=True)
            with open(path, "a", encoding="utf-8", errors="replace") as f:
                f.write(block)
    except OSError:
        pass


def install_app_error_logging() -> None:
    """Вызывать один раз при старте: лог для окон ошибок и необработанных исключений."""
    orig_showerror = messagebox.showerror

    def showerror_with_journal(title=None, message=None, **kw):
        append_error_journal(
            str(title if title is not None else "Ошибка"),
            str(message if message is not None else ""),
        )
        return orig_showerror(title, message, **kw)

    messagebox.showerror = showerror_with_journal  # type: ignore[method-assign]

    orig_excepthook = sys.__excepthook__

    def excepthook(exc_type, exc_value, exc_tb):
        try:
            tb_text = "".join(
                traceback.format_exception(exc_type, exc_value, exc_tb)
            )
            append_error_journal(
                "Необработанное исключение",
                str(exc_value),
                traceback_text=tb_text,
            )
        except Exception:
            pass
        orig_excepthook(exc_type, exc_value, exc_tb)

    sys.excepthook = excepthook
