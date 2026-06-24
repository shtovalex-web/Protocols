# -*- coding: utf-8 -*-
"""Сравнение версий major.minor[.patch] (1.5 трактуется как 1.5.0)."""

from __future__ import annotations

import re

_VERSION_RE = re.compile(r"^(\d+)\.(\d+)(?:\.(\d+))?")


def parse_version(value: str) -> tuple[int, int, int]:
    match = _VERSION_RE.match((value or "").strip())
    if not match:
        msg = f"Invalid version: {value!r}"
        raise ValueError(msg)
    patch = match.group(3)
    return int(match.group(1)), int(match.group(2)), int(patch) if patch else 0


def is_newer_version(candidate: str, current: str) -> bool:
    return parse_version(candidate) > parse_version(current)
