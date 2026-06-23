#!/usr/bin/env bash
# Копия комплекта сборки в ~/ProtocolOOT_linux_build (не с /mnt/* в WSL).
set -euo pipefail

TARGET="${PROTOCOLOOT_WSL_HOME:-$HOME/ProtocolOOT_linux_build}"
SRC="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

if [[ "$SRC" == /mnt/* ]]; then
  echo "Синхронизация: $SRC -> $TARGET"
  mkdir -p "$TARGET"
  if command -v rsync >/dev/null 2>&1; then
    rsync -a --delete \
      --exclude '.venv-linux/' \
      --exclude 'dist/' \
      --exclude 'release/_build/' \
      --exclude 'release/out_linux/' \
      --exclude 'release/out_linux.zip' \
      --exclude '__pycache__/' \
      --exclude 'app/__pycache__/' \
      "$SRC/" "$TARGET/"
  else
    rm -rf "$TARGET"
    mkdir -p "$TARGET"
    cp -a "$SRC/." "$TARGET/"
    rm -rf "$TARGET/.venv-linux" "$TARGET/dist" "$TARGET/release/_build" 2>/dev/null || true
  fi
  echo "Готово. Дальше:"
  echo "  cd $TARGET"
  echo "  chmod +x *.sh"
  echo "  ./check_env.sh && ./build.sh"
else
  TARGET="$SRC"
  echo "Каталог уже на локальной ФС: $TARGET"
  echo "Запуск: ./check_env.sh && ./build.sh"
fi
