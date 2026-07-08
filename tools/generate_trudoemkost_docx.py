# -*- coding: utf-8 -*-
"""
Сборка Word из Markdown: сравнение трудоёмкости формирования протоколов.

Исходник: docs/Трудоемкость_формирования_протоколов.md
Результат: docs/Трудоемкость_формирования_протоколов.docx

    py -3 tools/generate_trudoemkost_docx.py
"""

from __future__ import annotations

from pathlib import Path

from instruction_md_to_docx import build_docx_from_markdown

ROOT = Path(__file__).resolve().parent.parent
DOCS = ROOT / "docs"
MD_NAME = "Трудоемкость_формирования_протоколов.md"
DOCX_NAME = "Трудоемкость_формирования_протоколов.docx"


def main() -> None:
    build_docx_from_markdown(DOCS / MD_NAME, DOCS / DOCX_NAME)
    print(f"Записано: {DOCS / DOCX_NAME}")


if __name__ == "__main__":
    main()
