# -*- coding: utf-8 -*-
"""Починить CRLF в *.sh (если комплект копировали с Windows). Запуск: python3 fix_crlf.py"""

from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parent


def normalize_lf(path: Path) -> bool:
    data = path.read_bytes()
    if b"\r" not in data:
        return False
    text = data.decode("utf-8").replace("\r\n", "\n").replace("\r", "\n")
    path.write_text(text, encoding="utf-8", newline="\n")
    return True


def main() -> int:
    fixed = 0
    for sh in sorted(ROOT.glob("*.sh")):
        if normalize_lf(sh):
            print(f"исправлен: {sh.name}")
            fixed += 1
    if fixed:
        print(f"Готово: {fixed} файл(ов). Выполните: chmod +x *.sh && ./check_env.sh")
    else:
        print("CRLF не найдено в *.sh")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
