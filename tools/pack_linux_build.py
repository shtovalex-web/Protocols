# -*- coding: utf-8 -*-
"""
Сформировать автономный комплект для сборки на Linux (без всего репозитория).

  python tools/pack_linux_build.py
  python tools/pack_linux_build.py -o D:/temp/ProtocolOOT_linux_build

Результат: ProtocolOOT_linux_build/ — скопировать на Linux и выполнить:
  chmod +x *.sh && ./check_env.sh && ./build.sh
"""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
LINUX_PORT = ROOT / "linux_port"
DEFAULT_OUT = ROOT / "ProtocolOOT_linux_build"

if str(LINUX_PORT) not in sys.path:
    sys.path.insert(0, str(LINUX_PORT))

from sync_util import rmtree_resilient  # noqa: E402

SHELL_SCRIPTS = (
    "check_env.sh",
    "build.sh",
    "build_linux.sh",
    "install_deps.sh",
    "run.sh",
    "sync_workspace.sh",
)

COPY_FILES = (
    "build_linux.py",
    "verify_linux.py",
    "ruff_linux.py",
    "fix_crlf.py",
    "requirements.txt",
    "requirements-build.txt",
)

RELEASE_SKIP = {"_build", "out_linux", "out_linux.zip", "__pycache__"}


def _run_prepare() -> None:
    prepare = LINUX_PORT / "prepare.py"
    if not prepare.is_file():
        raise SystemExit(f"Нет {prepare}")
    rc = subprocess.run([sys.executable, str(prepare)], cwd=str(ROOT)).returncode
    if rc != 0:
        raise SystemExit(rc)


def _read_app_version(app_dir: Path) -> str:
    info = app_dir / "ProtocolOHT_next" / "protocol_app_info.py"
    if not info.is_file():
        return "unknown"
    text = info.read_text(encoding="utf-8")
    m = re.search(r'APP_VERSION\s*=\s*["\']([^"\']+)["\']', text)
    return m.group(1) if m else "unknown"


def _write_lf_text(src: Path, dst: Path) -> None:
    """Копировать текстовый файл с LF (важно для .sh на Linux)."""
    data = src.read_bytes()
    text = data.decode("utf-8").replace("\r\n", "\n").replace("\r", "\n")
    dst.parent.mkdir(parents=True, exist_ok=True)
    dst.write_text(text, encoding="utf-8", newline="\n")


def _copy_tree(src: Path, dst: Path, *, skip_names: set[str] | None = None) -> None:
    skip = skip_names or set()
    if not src.is_dir():
        return
    for item in src.rglob("*"):
        rel = item.relative_to(src)
        if any(part in skip for part in rel.parts):
            continue
        target = dst / rel
        if item.is_dir():
            target.mkdir(parents=True, exist_ok=True)
        else:
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(item, target)


def pack(output: Path, *, skip_prepare: bool = False) -> Path:
    if not skip_prepare:
        _run_prepare()

    app_src = LINUX_PORT / "app"
    release_src = LINUX_PORT / "release"
    if not (app_src / "main.py").is_file():
        raise SystemExit("Нет linux_port/app/main.py — prepare.py не выполнен.")
    if not (release_src / "build_release_linux.py").is_file():
        raise SystemExit("Нет linux_port/release/build_release_linux.py")

    if output.exists():
        rmtree_resilient(output)
    output.mkdir(parents=True)

    _copy_tree(app_src, output / "app")
    _copy_tree(release_src, output / "release", skip_names=RELEASE_SKIP)

    lib_src = LINUX_PORT / "lib"
    if lib_src.is_dir():
        _copy_tree(lib_src, output / "lib")

    for name in COPY_FILES:
        src = LINUX_PORT / name
        if src.is_file():
            shutil.copy2(src, output / name)

    for name in SHELL_SCRIPTS:
        src = LINUX_PORT / name
        if src.is_file():
            _write_lf_text(src, output / name)

    version = _read_app_version(app_src)
    stamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    (output / "VERSION.txt").write_text(
        f"ProtocolOOT build kit\nversion={version}\npacked_utc={stamp}\n",
        encoding="utf-8",
    )
    readme = release_src / "README_BUILD_LINUX.txt"
    if readme.is_file():
        shutil.copy2(readme, output / "README_BUILD_LINUX.txt")

    ruff_cfg = ROOT / "ruff.toml"
    if ruff_cfg.is_file():
        shutil.copy2(ruff_cfg, output / "ruff.toml")

    return output


def main() -> int:
    parser = argparse.ArgumentParser(description="Комплект ProtocolOOT_linux_build для сборки на Linux")
    parser.add_argument("-o", "--output", type=Path, default=DEFAULT_OUT, help="Папка выхода")
    parser.add_argument(
        "--skip-prepare",
        action="store_true",
        help="Не запускать prepare.py (если app/ уже актуален)",
    )
    args = parser.parse_args()

    out = pack(args.output.resolve(), skip_prepare=args.skip_prepare)
    print(f"Комплект сборки: {out}")
    print(f"  app/     — {sum(1 for _ in (out / 'app').rglob('*') if _.is_file())} файлов")
    print(f"  VERSION  — {(out / 'VERSION.txt').read_text(encoding='utf-8').strip()}")
    print()
    print("На Linux:")
    print(f"  cd {out.name}")
    print("  chmod +x *.sh")
    print("  ./install_deps.sh")
    print("  ./check_env.sh")
    print("  ./build.sh")
    print()
    print("Если ошибка bash\\r: python3 fix_crlf.py && chmod +x *.sh")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
