#!/usr/bin/env bash
# Запуск приложения из app/ (исходники, не бинарник).
set -euo pipefail

KIT_ROOT="$(cd "$(dirname "$0")" && pwd)"
APP="$KIT_ROOT/app"
# shellcheck source=lib/sh_common.sh
source "$KIT_ROOT/lib/sh_common.sh"

BUILD_PYTHON="$(resolve_build_python || true)"

if [[ ! -f "$APP/main.py" ]]; then
  echo "Нет $APP/main.py" >&2
  exit 1
fi

if ! "$BUILD_PYTHON" -c "import tkinter" 2>/dev/null; then
  if is_altlinux; then
    echo "Нужен tkinter: sudo apt install python3-modules-tkinter python3.11" >&2
  else
    echo "Нужен tkinter: sudo apt install python3-tk" >&2
  fi
  exit 1
fi

if [[ -f "$KIT_ROOT/.venv-linux/bin/activate" ]]; then
  ensure_venv "$KIT_ROOT" "$BUILD_PYTHON"
fi

cd "$APP"
exec "$BUILD_PYTHON" main.py
