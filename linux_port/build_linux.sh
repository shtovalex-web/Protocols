#!/usr/bin/env bash
# Сборка Linux-бинарника (PyInstaller) и комплекта ProtocolOHT_linux_dist/.
set -euo pipefail
ROOT="$(cd "$(dirname "$0")" && pwd)"

if ! python3 -c "import tkinter" 2>/dev/null; then
  echo "Пакет python3-tk не найден. Установите: sudo apt install python3-tk" >&2
  exit 1
fi

if ! command -v objdump >/dev/null 2>&1; then
  echo "Для PyInstaller нужен objdump (пакет binutils): sudo apt install binutils" >&2
  exit 1
fi

if ! ldconfig -p 2>/dev/null | grep -q 'libpython3.*\.so' && \
   ! find /usr/lib64 /usr/lib -name 'libpython3*.so*' 2>/dev/null | grep -q .; then
  echo "Не найдена libpython (нужна для PyInstaller). Установите, например:" >&2
  echo "  sudo apt install python3-dev" >&2
  exit 1
fi

if [[ -f "$ROOT/.venv/bin/activate" ]]; then
  # shellcheck disable=SC1091
  source "$ROOT/.venv/bin/activate"
fi

python3 -m pip install -r "$ROOT/requirements-build.txt"
python3 "$ROOT/build_linux.py" "$@"
