#!/usr/bin/env bash
# Сборка ProtocolOOT: проверка окружения → venv → PyInstaller → release/out_linux/
set -euo pipefail

KIT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$KIT_ROOT"
# shellcheck source=lib/sh_common.sh
source "$KIT_ROOT/lib/sh_common.sh"

VENV_DIR="$KIT_ROOT/.venv-linux"
RELEASE_SCRIPT="$KIT_ROOT/release/build_release_linux.py"
BUILD_PYTHON="$(resolve_build_python || true)"

SKIP_CHECK=0
BUILD_ARGS=()
for arg in "$@"; do
  if [[ "$arg" == "--skip-check" ]]; then
    SKIP_CHECK=1
  else
    BUILD_ARGS+=("$arg")
  fi
done

if [[ "$SKIP_CHECK" -eq 0 ]]; then
  bash "$KIT_ROOT/check_env.sh"
fi

if [[ ! -f "$RELEASE_SCRIPT" ]]; then
  echo "Не найден $RELEASE_SCRIPT" >&2
  exit 1
fi

if ! "$BUILD_PYTHON" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)' 2>/dev/null; then
  echo "Нужен Python 3.10+. На ALT: sudo apt install python3.11 python3.11-tools python3.11-devel libpython3.11" >&2
  exit 1
fi

ensure_venv "$KIT_ROOT" "$BUILD_PYTHON"
python -m pip install -q --upgrade pip
python -m pip install -q -r "$KIT_ROOT/requirements-build.txt"
python "$RELEASE_SCRIPT" --local "${BUILD_ARGS[@]}"
