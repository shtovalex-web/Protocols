# -*- coding: utf-8 -*-
"""Tcl/Tk для PyInstaller --onefile: сброс унаследованных путей и _MEIPASS."""

from __future__ import annotations

import os
import sys

_TK_ENV_KEYS = ("TCL_LIBRARY", "TK_LIBRARY", "TCLLIBPATH", "TKPATH")


def configure_frozen_tk_environment(
    *,
    frozen: bool | None = None,
    meipass: str | None = None,
) -> None:
    is_frozen = frozen if frozen is not None else getattr(sys, "frozen", False)
    if not is_frozen:
        return
    base = meipass if meipass is not None else getattr(sys, "_MEIPASS", "")
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
