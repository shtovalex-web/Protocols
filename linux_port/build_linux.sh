#!/usr/bin/env bash
# Обёртка: ./build_linux.sh → check_env.sh + build.sh (совместимость с документацией).
set -euo pipefail
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
exec bash "$DIR/build.sh" "$@"
