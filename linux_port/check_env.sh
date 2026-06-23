#!/usr/bin/env bash
# Пошаговая проверка окружения перед сборкой ProtocolOOT (Linux).
set -uo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
KIT_ROOT="$SCRIPT_DIR"
APP_DIR="$KIT_ROOT/app"
RELEASE_SCRIPT="$KIT_ROOT/release/build_release_linux.py"
VENV_DIR="$KIT_ROOT/.venv-linux"

# shellcheck source=lib/sh_common.sh
source "$KIT_ROOT/lib/sh_common.sh"

BUILD_PYTHON="$(resolve_build_python || true)"
DEFAULT_PYTHON="python3"
command -v python3 >/dev/null 2>&1 || DEFAULT_PYTHON=""

FAILS=0
STEP=0
TOTAL=17

_ok() {
  STEP=$((STEP + 1))
  printf '[%02d/%02d] %s — OK%s\n' "$STEP" "$TOTAL" "$1" "${2:+ ($2)}"
}

_fail() {
  STEP=$((STEP + 1))
  FAILS=$((FAILS + 1))
  printf '[%02d/%02d] %s — FAIL\n' "$STEP" "$TOTAL" "$1" >&2
  while [[ $# -gt 1 ]]; do
    shift
    printf '         %s\n' "$1" >&2
  done
}

_skip() {
  STEP=$((STEP + 1))
  printf '[%02d/%02d] %s — пропуск (%s)\n' "$STEP" "$TOTAL" "$1" "$2"
}

_alt_python_hint() {
  echo "ALT Linux: sudo apt install python3.11 python3.11-tools python3.11-dev libpython3.11 python3-module-pip python3-modules-tkinter"
  echo "Затем: rm -rf .venv-linux && ./install_deps.sh"
}

echo "=== Проверка окружения сборки ProtocolOOT ==="
echo "Каталог комплекта: $KIT_ROOT"
if is_altlinux; then
  echo "Дистрибутив: ALT Linux"
fi
echo "Python для сборки: ${BUILD_PYTHON:-?} ($("$BUILD_PYTHON" --version 2>&1 || echo 'не найден'))"
echo

# 1. Путь
case "$KIT_ROOT" in
  /mnt/*|/media/*)
    _fail "Путь к проекту" \
      "Сборка на /mnt/* часто ломает venv/PyInstaller." \
      "Скопируйте комплект в ~/ProtocolOOT_linux_build"
    ;;
  *)
    _ok "Путь к проекту" "$KIT_ROOT"
    ;;
esac

# 2. python3 в PATH
if [[ -n "$DEFAULT_PYTHON" ]]; then
  PY_VER="$("$DEFAULT_PYTHON" -c 'import sys; print(".".join(map(str, sys.version_info[:3])))' 2>/dev/null || echo "?")"
  _ok "Команда python3" "$PY_VER"
else
  _fail "Команда python3" "Установите Python 3.10+"
fi

# 3. Версия для сборки (>= 3.10)
if "$BUILD_PYTHON" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)' 2>/dev/null; then
  if "$BUILD_PYTHON" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 11) else 1)' 2>/dev/null; then
    _ok "Версия Python (сборка)" ">= 3.11 ($BUILD_PYTHON)"
  else
    _ok "Версия Python (сборка)" ">= 3.10 ($BUILD_PYTHON)"
  fi
else
  if is_altlinux; then
    _fail "Версия Python (сборка)" \
      "Сейчас: $(python3 --version 2>&1). Нужен python3.11:" \
      "sudo apt install python3.11 python3.11-tools python3.11-dev libpython3.11"
  else
    _fail "Версия Python (сборка)" \
      "Нужен Python 3.10+, сейчас: $(python3 --version 2>&1)" \
      "Debian/Ubuntu: sudo apt install python3.11 python3.11-venv python3.11-dev"
  fi
fi

# 4. venv
if "$BUILD_PYTHON" -m venv --help >/dev/null 2>&1; then
  _ok "Модуль venv ($BUILD_PYTHON)"
elif is_altlinux; then
  _fail "Модуль venv" "sudo apt install python3.11-tools"
else
  _fail "Модуль venv" "sudo apt install python3-venv"
fi

# 5. tkinter (для того же Python, что пойдёт в venv)
if "$BUILD_PYTHON" -c "import tkinter" 2>/dev/null; then
  _ok "tkinter ($BUILD_PYTHON)"
elif is_altlinux; then
  _fail "tkinter" "sudo apt install python3-modules-tkinter (и python3.11 для сборки)"
else
  _fail "tkinter" "sudo apt install python3-tk"
fi

# 6. objdump
if command -v objdump >/dev/null 2>&1; then
  _ok "objdump (binutils)" "$(objdump --version 2>/dev/null | head -n1 || true)"
else
  _fail "objdump (binutils)" "sudo apt install binutils"
fi

# 7. libpython для BUILD_PYTHON
if has_libpython_for "$BUILD_PYTHON"; then
  LP_DETAIL="$("$BUILD_PYTHON" "$KIT_ROOT/lib/libpython_probe.py" --verbose 2>/dev/null || true)"
  _ok "libpython ($BUILD_PYTHON)" "${LP_DETAIL:-$(python_short_version "$BUILD_PYTHON")}"
elif [[ ! -f "$VENV_DIR/bin/python" ]]; then
  _skip "libpython ($BUILD_PYTHON)" "после ./install_deps.sh"
else
  if is_altlinux; then
    _fail "libpython" \
      "sudo apt-get install -y libpython3.11 python3.11-dev" \
      "rm -rf .venv-linux && ./install_deps.sh && ./check_env.sh"
  else
    _fail "libpython" \
      "sudo apt-get install -y python3-dev (или python3.11-dev)" \
      "./install_deps.sh && ./check_env.sh"
  fi
fi

# 8. Место на диске
AVAIL_KB="$(df -Pk "$KIT_ROOT" 2>/dev/null | awk 'NR==2 {print $4}' || echo 0)"
if [[ "${AVAIL_KB:-0}" =~ ^[0-9]+$ ]] && (( AVAIL_KB >= 2097152 )); then
  _ok "Место на диске" "$(( AVAIL_KB / 1024 / 1024 )) ГБ"
elif [[ "${AVAIL_KB:-0}" =~ ^[0-9]+$ ]] && (( AVAIL_KB >= 524288 )); then
  STEP=$((STEP + 1))
  printf '[%02d/%02d] Место на диске — OK (%s МБ)\n' "$STEP" "$TOTAL" "$(( AVAIL_KB / 1024 ))"
else
  _fail "Место на диске" "Нужно >= 512 МБ"
fi

# 9–11. Файлы
if [[ -f "$APP_DIR/main.py" ]]; then _ok "app/main.py"; else _fail "app/main.py" "pack_linux_build.py или prepare.py"; fi
if [[ -f "$RELEASE_SCRIPT" ]]; then _ok "release/build_release_linux.py"; else _fail "release/build_release_linux.py" "git pull / pack"; fi
if [[ -f "$KIT_ROOT/release/protocol_oot_linux.spec" ]]; then _ok "release/protocol_oot_linux.spec"; else _fail "release/protocol_oot_linux.spec" "нет в комплекте"; fi

# 12. CRLF
CRLF_FOUND=""
for sh in "$KIT_ROOT"/*.sh; do
  [[ -f "$sh" ]] || continue
  if grep -q $'\r' "$sh" 2>/dev/null; then CRLF_FOUND="$(basename "$sh")"; break; fi
done
if [[ -n "$CRLF_FOUND" ]]; then
  _fail "Shell-скрипты (LF)" "sed -i 's/\\r\$//' *.sh"
else
  _ok "Shell-скрипты (LF)"
fi

# 13–14. requirements и lib
if [[ -f "$KIT_ROOT/requirements-build.txt" && -f "$KIT_ROOT/requirements.txt" ]]; then
  _ok "requirements.txt + requirements-build.txt"
else
  _fail "requirements" "нет requirements*.txt"
fi
if [[ -f "$KIT_ROOT/lib/sh_common.sh" ]]; then
  _ok "lib/sh_common.sh"
else
  _fail "lib/sh_common.sh" "Обновите комплект (pack_linux_build.py)"
fi

# 15–17. venv
if [[ -f "$VENV_DIR/bin/python" ]]; then
  if "$VENV_DIR/bin/python" -m pip --version >/dev/null 2>&1; then
    _ok "venv .venv-linux" "$("$VENV_DIR/bin/python" --version 2>&1)"
  else
    _fail "venv .venv-linux" "rm -rf .venv-linux && ./install_deps.sh"
  fi
  if "$VENV_DIR/bin/python" - <<'PY'
import importlib
for mod in ("openpyxl", "docx", "fpdf", "pymorphy2"):
    importlib.import_module(mod)
PY
  then
    _ok "pip-зависимости"
  else
    _fail "pip-зависимости" "./install_deps.sh"
  fi
  if "$VENV_DIR/bin/python" -c "import PyInstaller" 2>/dev/null; then
    _ok "PyInstaller" "$("$VENV_DIR/bin/python" -m PyInstaller --version 2>/dev/null || echo "?")"
  else
    _fail "PyInstaller" "pip install -r requirements-build.txt"
  fi
else
  _skip "venv" "создастся install_deps.sh"
  _skip "pip-зависимости" "нет venv"
  _skip "PyInstaller" "нет venv"
fi

echo
if (( FAILS > 0 )); then
  echo "Итог: $FAILS ошибок — устраните FAIL, затем ./build.sh" >&2
  if is_altlinux && ! "$BUILD_PYTHON" -c 'import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)' 2>/dev/null; then
    echo >&2
    _alt_python_hint >&2
  fi
  exit 1
fi
echo "Итог: проверки пройдены."
if [[ ! -f "$VENV_DIR/bin/python" ]]; then
  echo "Дальше: ./install_deps.sh && ./build.sh"
else
  echo "Запуск: ./build.sh"
fi
exit 0
