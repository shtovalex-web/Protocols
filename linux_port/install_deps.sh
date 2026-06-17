#!/usr/bin/env bash
# Установка зависимостей Linux-порта (системные пакеты + pip).
set -euo pipefail
ROOT="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT"

if ! command -v python3 >/dev/null 2>&1; then
  echo "Нужен python3 (3.10+)." >&2
  exit 1
fi

if command -v apt-get >/dev/null 2>&1; then
  echo "Системные пакеты (Debian/Ubuntu)..."
  sudo apt-get update
  sudo apt-get install -y \
    python3-tk \
    python3-venv \
    python3-dev \
    binutils \
    libreoffice-writer \
    fonts-dejavu-core \
    fonts-liberation
fi

if [[ ! -d "$ROOT/.venv" ]]; then
  python3 -m venv "$ROOT/.venv"
fi
# shellcheck disable=SC1091
source "$ROOT/.venv/bin/activate"
python -m pip install --upgrade pip
python -m pip install -r "$ROOT/requirements.txt"

if [[ ! -f "$ROOT/app/main.py" ]]; then
  echo "Копия app/ не найдена — выполните: python3 prepare.py" >&2
  exit 1
fi

echo "Готово. Запуск: ./run.sh"
