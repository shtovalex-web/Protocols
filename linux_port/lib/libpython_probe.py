#!/usr/bin/env python3
"""Проверка libpython для PyInstaller: exit 0 — можно собирать, 1 — нет."""
from __future__ import annotations

import os
import sys
import sysconfig


def _is_file(path: str | None) -> bool:
    return bool(path) and os.path.isfile(path)


def libpython_ready() -> tuple[bool, str]:
    libdir = sysconfig.get_config_var("LIBDIR") or ""
    version = sysconfig.get_config_var("VERSION") or (
        f"{sys.version_info.major}.{sys.version_info.minor}"
    )
    multiarch = sysconfig.get_config_var("MULTIARCH") or ""
    prefix = sysconfig.get_config_var("prefix") or ""

    candidates: list[str] = []
    for name in (
        sysconfig.get_config_var("INSTSONAME"),
        sysconfig.get_config_var("LDLIBRARY"),
        sysconfig.get_config_var("LIBRARY"),
    ):
        if not name:
            continue
        if os.path.isabs(name):
            candidates.append(name)
        else:
            candidates.append(os.path.join(libdir, name))

    libpl = sysconfig.get_config_var("LIBPL") or ""
    if libpl and os.path.isdir(libpl):
        try:
            for entry in os.listdir(libpl):
                if entry.startswith("libpython") and (".so" in entry or entry.endswith(".a")):
                    candidates.append(os.path.join(libpl, entry))
        except OSError:
            pass

    for base in (f"libpython{version}", f"libpython{version}m"):
        for suffix in (".so", ".so.1.0", ".a"):
            candidates.append(os.path.join(libdir, base + suffix))
            if multiarch:
                candidates.append(os.path.join(libdir, multiarch, base + suffix))
        for sub in ("lib64", "lib"):
            candidates.append(os.path.join(prefix, sub, base + ".so.1.0"))
            candidates.append(os.path.join(prefix, sub, f"python{version}", base + ".so.1.0"))

    seen: set[str] = set()
    for path in candidates:
        if path in seen:
            continue
        seen.add(path)
        if _is_file(path):
            return True, path

    # Статический интерпретатор + заголовки (python3.11-dev на ALT)
    if not sysconfig.get_config_var("Py_ENABLE_SHARED"):
        libpc = sysconfig.get_config_var("LIBPC") or ""
        header = os.path.join(libpc, "Python.h")
        if _is_file(header):
            return True, "static+Python.h"

    return False, ""


if __name__ == "__main__":
    ok, detail = libpython_ready()
    if ok and "--verbose" in sys.argv:
        print(detail)
    sys.exit(0 if ok else 1)
