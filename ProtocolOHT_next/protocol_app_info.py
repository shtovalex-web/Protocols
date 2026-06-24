# -*- coding: utf-8 -*-
"""Сведения о приложении для окна «О программе» и заголовка окна.

Логотип и значок окон — вшитые PNG в protocol_embedded_assets.py (перегенерация: tools/_gen_embedded_png.py).
"""

from __future__ import annotations

import re
import webbrowser
from urllib.parse import quote

import tkinter as tk

from ui_theme import Colors

# Коротко — в заголовке окна и панели задач (не обрезается).
APP_WINDOW_TITLE = "ProtocolOOT — протоколы ОТ"
# Полное название — в «О программе» и документации.
APP_FULL_NAME = "Программа для формирования протоколов проверки знаний по охране труда"
# Отредактируйте перед распространением сборки.
# При смене версии дополните bundle/ЖУРНАЛ_ДОРАБОТОК.md
APP_VERSION = "1.6.3"
APP_DEVELOPER_NAME = "Шитов Алексей Александрович"
APP_DEVELOPER_ROLE = "разработка и сопровождение"
APP_DEVELOPER_ORG = ""
APP_DEVELOPER_CONTACT = "ShitovAA@Transneft.ru"
# Тема письма при клике на e-mail в окне «О программе».
APP_FEEDBACK_MAIL_SUBJECT = (
    "Предложение/замечание по работе с программой протоколов"
)

_ABOUT_EMAIL_TAG = "about_email_link"
_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", re.IGNORECASE)


def _is_email_address(value: str) -> bool:
    return bool(_EMAIL_RE.match((value or "").strip()))


def feedback_mailto_url(email: str) -> str:
    """mailto: с темой для обратной связи по программе."""
    addr = (email or "").strip()
    subject = quote((APP_FEEDBACK_MAIL_SUBJECT or "").strip(), safe="")
    if subject:
        return f"mailto:{addr}?subject={subject}"
    return f"mailto:{addr}"


def _about_plain_lines() -> list[str]:
    """Строки текста «О программе» (без разметки ссылок)."""
    lines: list[str] = [APP_FULL_NAME]
    ver = (APP_VERSION or "").strip()
    if ver:
        lines.append(f"Версия: {ver}")
    lines.append("")
    blocks: list[str] = []
    name = (APP_DEVELOPER_NAME or "").strip()
    if name:
        blocks.append(f"Разработчик: {name}")
    role = (APP_DEVELOPER_ROLE or "").strip()
    if role:
        blocks.append(role)
    org = (APP_DEVELOPER_ORG or "").strip()
    if org:
        blocks.append(f"Организация: {org}")
    contact = (APP_DEVELOPER_CONTACT or "").strip()
    if contact:
        blocks.append(f"Контакт: {contact}")
    if blocks:
        lines.extend(blocks)
    else:
        lines.append(
            "Сведения о разработчике не заданы.\n"
            "Их можно указать в файле protocol_app_info.py (константы APP_VERSION, APP_DEVELOPER_*)."
        )
    return lines


def format_application_about_text() -> str:
    """Текст окна «О программе» (простой текст)."""
    return "\n".join(_about_plain_lines())


def populate_application_about_text(widget: tk.Text) -> None:
    """Заполняет виджет «О программе»; e-mail — кликабельная ссылка mailto:."""
    widget.configure(state="normal")
    widget.delete("1.0", "end")
    contact = (APP_DEVELOPER_CONTACT or "").strip()
    email_link = contact if _is_email_address(contact) else ""

    for line in _about_plain_lines():
        if email_link and line == f"Контакт: {contact}":
            widget.insert("end", "Контакт: ")
            widget.insert("end", contact, (_ABOUT_EMAIL_TAG,))
            widget.insert("end", "\n")
            continue
        widget.insert("end", line + "\n")

    if not email_link:
        return

    widget.tag_configure(_ABOUT_EMAIL_TAG, foreground=Colors.link, underline=True)

    def _open_mail(_event=None) -> str:
        webbrowser.open(feedback_mailto_url(email_link))
        return "break"

    widget.tag_bind(_ABOUT_EMAIL_TAG, "<Button-1>", _open_mail)
    widget.tag_bind(_ABOUT_EMAIL_TAG, "<Enter>", lambda _e: widget.configure(cursor="hand2"))
    widget.tag_bind(_ABOUT_EMAIL_TAG, "<Leave>", lambda _e: widget.configure(cursor=""))

    def _click_at_pointer(event) -> str | None:
        try:
            index = widget.index(f"@{event.x},{event.y}")
        except tk.TclError:
            return None
        if _ABOUT_EMAIL_TAG in widget.tag_names(index):
            return _open_mail()
        return None

    widget.bind("<Button-1>", _click_at_pointer, add="+")
