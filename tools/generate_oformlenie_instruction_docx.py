# -*- coding: utf-8 -*-
"""
Сборка Word из Markdown (постоянный цикл правок).

Исходник: bundle/ИНСТРУКЦИЯ_оформление_протоколов_Минтруд.md
Результат: bundle/ИНСТРУКЦИЯ_оформление_протоколов_Минтруд.docx

    py -3 tools/generate_oformlenie_instruction_docx.py

Или: generate_oformlenie_instruction_docx.bat
"""

from __future__ import annotations

from pathlib import Path

from instruction_md_to_docx import build_docx_from_markdown

ROOT = Path(__file__).resolve().parent.parent
BUNDLE = ROOT / "bundle"
MD_NAME = "ИНСТРУКЦИЯ_оформление_протоколов_Минтруд.md"
DOCX_NAME = "ИНСТРУКЦИЯ_оформление_протоколов_Минтруд.docx"


def main() -> None:
    build_docx_from_markdown(BUNDLE / MD_NAME, BUNDLE / DOCX_NAME)
    print(f"Записано: {BUNDLE / DOCX_NAME}")


if __name__ == "__main__":
    main()
