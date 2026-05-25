# -*- coding: utf-8 -*-
"""Включить защиту «только чтение» в Word для стандартных шаблонов. py -3 tools/protect_protocol_templates.py"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from docx_template_protection import protect_standard_protocol_templates


def main() -> int:
    done = protect_standard_protocol_templates(ROOT)
    if not done:
        print("Шаблоны не найдены.")
        return 1
    for line in done:
        print(line)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
