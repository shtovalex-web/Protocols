# -*- coding: utf-8 -*-
"""PyInstaller runtime hook: Tcl/Tk до import tkinter.

Должен быть самодостаточным — без импортов из проекта (в onefile их нет в sys.path).
"""

from __future__ import annotations

import os
import sys

_TK_ENV_KEYS = ("TCL_LIBRARY", "TK_LIBRARY", "TCLLIBPATH", "TKPATH")


def _configure_frozen_tk_environment() -> None:
    if not getattr(sys, "frozen", False):
        return
    base = getattr(sys, "_MEIPASS", "")
    if not base:
        return
    for key in _TK_ENV_KEYS:
        os.environ.pop(key, None)
    tcl_dir = os.path.join(base, "_tcl_data")
    tk_dir = os.path.join(base, "_tk_data")
    if os.path.isfile(os.path.join(tcl_dir, "init.tcl")):
        os.environ["TCL_LIBRARY"] = tcl_dir
    if os.path.isdir(tk_dir):
        os.environ["TK_LIBRARY"] = tk_dir


_configure_frozen_tk_environment()
