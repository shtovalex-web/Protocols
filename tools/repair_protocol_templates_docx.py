# -*- coding: utf-8 -*-
"""
Восстановление шаблонов .docx после повреждения settings.xml.

Закройте Word. Из корня: py -3 tools/repair_protocol_templates_docx.py
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

BACKUP_BUNDLE = ROOT / "ib_full_no_data_20260429_214543" / "bundle"
NAMES = ("default_protocol.docx", "default_protocol_tehnicheskiy.docx")


def _restore_from_backup() -> list[str]:
    done: list[str] = []
    if not BACKUP_BUNDLE.is_dir():
        return done
    from docx_template_protection import set_windows_readonly_flag

    for name in NAMES:
        src = BACKUP_BUNDLE / name
        if not src.is_file():
            continue
        for dest in (ROOT / name, ROOT / "bundle" / name):
            dest.parent.mkdir(parents=True, exist_ok=True)
            if dest.is_file():
                set_windows_readonly_flag(dest, False)
            shutil.copy2(src, dest)
            done.append(str(dest))
    return done


def main() -> int:
    from docx_template_protection import unprotect_standard_protocol_templates

    print("Снятие защиты…")
    ok, err = unprotect_standard_protocol_templates(ROOT)
    for line in ok:
        print(" ", line)
    for line in err:
        print("  ОШИБКА:", line)

    print("Копирование эталона из резервной папки…")
    restored = _restore_from_backup()
    if not restored:
        print(
            "  Резерв ib_full_no_data_20260429_214543/bundle не найден — "
            "используются текущие файлы после снятия защиты."
        )
    for line in restored:
        print(" ", line)

    print("Маркеры V_PROF и защита Word…")
    import runpy

    runpy.run_path(str(ROOT / "tools" / "patch_protocol_template_markers.py"), run_name="__main__")
    print("Готово. Откройте шаблон в Word для проверки.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
