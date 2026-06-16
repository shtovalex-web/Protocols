#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")" && pwd)"
APP="$ROOT/app"

if [[ ! -f "$APP/main.py" ]]; then
  echo "Нет $APP/main.py — выполните: python3 prepare.py" >&2
  exit 1
fi

if [[ -f "$ROOT/.venv/bin/activate" ]]; then
  # shellcheck disable=SC1091
  source "$ROOT/.venv/bin/activate"
fi

cd "$APP"
exec python3 main.py
