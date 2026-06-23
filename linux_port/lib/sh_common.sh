#!/usr/bin/env bash
# Общие функции для check_env.sh, install_deps.sh, build.sh, run.sh

is_altlinux() {
  [[ -f /etc/os-release ]] || return 1
  # shellcheck disable=SC1091
  . /etc/os-release
  [[ "${ID:-}" == "altlinux" || "${ID_LIKE:-}" == *alt* ]]
}

# Интерпретатор для сборки: python3.12 → 3.11 → 3.10 → python3 (>= 3.10).
resolve_build_python() {
  local c
  for c in python3.12 python3.11 python3.10 python3; do
    command -v "$c" >/dev/null 2>&1 || continue
    if "$c" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)' 2>/dev/null; then
      echo "$c"
      return 0
    fi
  done
  echo python3
  return 1
}

python_short_version() {
  local py="${1:-python3}"
  "$py" -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")'
}

has_libpython_for() {
  local py="${1:-python3}"
  local ver
  ver="$(python_short_version "$py" 2>/dev/null || echo 3)"
  if ldconfig -p 2>/dev/null | grep -qE "libpython${ver//./\\.}.*\\.so"; then
    return 0
  fi
  if find /usr/lib64 /usr/lib -name "libpython${ver}*.so*" 2>/dev/null | grep -q .; then
    return 0
  fi
  if find /usr/lib64 /usr/lib -name 'libpython3*.so*' 2>/dev/null | grep -q .; then
    return 0
  fi
  return 1
}

install_system_packages() {
  if ! command -v apt-get >/dev/null 2>&1; then
    echo "apt-get не найден — установите пакеты вручную (см. README_BUILD_LINUX.txt)." >&2
    return 0
  fi
  sudo apt-get update
  if is_altlinux; then
    echo "Системные пакеты (ALT Linux)..."
    sudo apt-get install -y \
      python3.11 \
      python3.11-tools \
      python3-module-pip \
      libpython3.11 \
      python3.11-devel \
      python3-modules-tkinter \
      binutils \
      xvfb \
      libreoffice \
      fonts-dejavu \
      fonts-ttf-dejavu
  else
    echo "Системные пакеты (Debian/Ubuntu)..."
    sudo apt-get install -y \
      python3-tk \
      python3-venv \
      python3-dev \
      binutils \
      xvfb \
      libreoffice-writer \
      fonts-dejavu-core \
      fonts-liberation
  fi
}

ensure_venv() {
  local kit_root="$1"
  local py="$2"
  local venv_dir="$kit_root/.venv-linux"
  local want_ver
  want_ver="$(python_short_version "$py")"
  if [[ -f "$venv_dir/pyvenv.cfg" ]]; then
    if ! grep -q "version = ${want_ver}" "$venv_dir/pyvenv.cfg" 2>/dev/null; then
      echo "Пересоздание .venv-linux (нужен Python ${want_ver}, был другой)..."
      rm -rf "$venv_dir"
    fi
  fi
  if [[ ! -f "$venv_dir/bin/activate" ]]; then
    "$py" -m venv "$venv_dir"
    "$venv_dir/bin/python" -m ensurepip --upgrade >/dev/null 2>&1 || true
  fi
  # shellcheck disable=SC1091
  source "$venv_dir/bin/activate"
}
