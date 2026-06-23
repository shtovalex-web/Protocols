# -*- coding: utf-8 -*-
"""Общий запуск ruff для Linux-копии (тот же конфиг, что verify_project на Windows)."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def ruff_config_path(linux_port: Path, app_dir: Path) -> Path | None:
    roots = [linux_port]
    if linux_port.name == "linux_port":
        roots.append(linux_port.parent)
    for root in roots:
        candidate = root / "ruff.toml"
        if candidate.is_file():
            return candidate.resolve()
    app_cfg = app_dir / "ruff.toml"
    if app_cfg.is_file():
        return app_cfg.resolve()
    return None


def run_ruff_on_app(
    *,
    linux_port: Path,
    app_dir: Path,
    cwd: Path | None = None,
    capture: bool = False,
) -> subprocess.CompletedProcess[str] | int:
    cmd = [sys.executable, "-m", "ruff", "check", str(app_dir)]
    cfg = ruff_config_path(linux_port, app_dir)
    if cfg is not None:
        cmd.extend(["--config", str(cfg)])
    if capture:
        return subprocess.run(
            cmd,
            cwd=str(cwd or linux_port),
            capture_output=True,
            text=True,
        )
    return subprocess.run(cmd, cwd=str(cwd or linux_port)).returncode
