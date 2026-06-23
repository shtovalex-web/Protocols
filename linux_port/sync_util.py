# -*- coding: utf-8 -*-
"""Общие утилиты синхронизации Linux-копии (prepare, pack, hooks)."""

from __future__ import annotations

import os
import shutil
import stat
from pathlib import Path

# Источники для linux_port/app/ (должны совпадать с prepare.py).
COPY_DIRS = (
    "ProtocolOHT_next",
    "bundle",
)

COPY_ROOT_FILES = (
    "main.py",
    "app_paths.py",
    "clipboard_ui.py",
    "commission_admin.py",
    "docx_template_protection.py",
    "employees_io.py",
    "excel_data_cache.py",
    "faq_viewer.py",
    "mintrud_export.py",
    "mintrud_trained_registry.py",
    "program_keys.py",
    "programs_v_prof.py",
    "russian_genitive.py",
    "v_prof_combinations.py",
    "v_program_registry_match.py",
    "educated_person_import_v1.0.9.xsd",
)

_LINUX_OVERLAY_PREFIX = "linux_port/overlays/"


def path_affects_linux_app(rel_path: str) -> bool:
    """Нужно ли обновлять linux_port/app/ после изменения файла в репозитории."""
    p = (rel_path or "").replace("\\", "/").lstrip("./")
    if not p or p.startswith("linux_port/app/"):
        return False
    if p in COPY_ROOT_FILES:
        return True
    for dirname in COPY_DIRS:
        if p == dirname or p.startswith(f"{dirname}/"):
            return True
    if p.startswith(_LINUX_OVERLAY_PREFIX):
        return True
    if p == "linux_port/prepare.py":
        return True
    return False


def rmtree_resilient(path: Path) -> None:
    """Удалить каталог; на Windows снимает read-only при PermissionError."""

    def onexc(func, p, exc_info):
        exc = exc_info[1] if isinstance(exc_info, tuple) else exc_info
        if isinstance(exc, PermissionError):
            os.chmod(p, stat.S_IWRITE)
            func(p)
            return
        raise exc

    shutil.rmtree(path, onexc=onexc)
