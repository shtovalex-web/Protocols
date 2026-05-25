# -*- coding: utf-8 -*-
"""
Сборка Word из bundle/ПОДРОБНАЯ_ИНСТРУКЦИЯ_для_пользователя.md

    py -3 tools/generate_podrobnaya_instruction_docx.py

Или: generate_podrobnaya_instruction_docx.bat
"""

from __future__ import annotations

from pathlib import Path

from instruction_md_to_docx import build_docx_from_markdown

ROOT = Path(__file__).resolve().parent.parent
BUNDLE = ROOT / "bundle"
MD_NAME = "ПОДРОБНАЯ_ИНСТРУКЦИЯ_для_пользователя.md"
DOCX_NAME = "ПОДРОБНАЯ_ИНСТРУКЦИЯ_для_пользователя.docx"


def main() -> None:
    md_path = BUNDLE / MD_NAME
    out_path = BUNDLE / DOCX_NAME
    build_docx_from_markdown(md_path, out_path)
    print(f"Записано: {out_path}")


if __name__ == "__main__":
    main()
