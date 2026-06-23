#!/usr/bin/env bash
# Установка зависимостей для сборки/запуска (системные пакеты + .venv-linux).
set -euo pipefail

KIT_ROOT="$(cd "$(dirname "$0")" && pwd)"
cd "$KIT_ROOT"
# shellcheck source=lib/sh_common.sh
source "$KIT_ROOT/lib/sh_common.sh"

BUILD_PYTHON="$(resolve_build_python || true)"
if ! "$BUILD_PYTHON" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)' 2>/dev/null; then
  echo "Нужен Python 3.10+ (для ALT: sudo apt install python3.11 python3.11-tools python3.11-dev libpython3.11)." >&2
  exit 1
fi
echo "Python для сборки: $BUILD_PYTHON ($("$BUILD_PYTHON" --version 2>&1))"

install_system_packages

if [[ ! -f "$KIT_ROOT/app/main.py" ]]; then
  echo "Нет app/main.py в комплекте." >&2
  exit 1
fi

ensure_venv "$KIT_ROOT" "$BUILD_PYTHON"
python -m pip install --upgrade pip
python -m pip install -r "$KIT_ROOT/requirements-build.txt"

echo
echo "Готово. Дальше:"
echo "  ./check_env.sh"
echo "  ./build.sh"
