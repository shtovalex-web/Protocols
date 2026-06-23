# Сборка ProtocolOOT (Linux)

## Автономный комплект (без всего репозитория)

На Windows (перед переносом на Linux):

```text
python linux_port/prepare.py
python tools/pack_linux_build.py
```

Получите папку `ProtocolOOT_linux_build/` (~30–80 МБ) — скопируйте или заархивируйте на Linux-ПК.

## На Linux-ПК

```bash
cd ProtocolOOT_linux_build   # или linux_port/ из git clone -b linux
chmod +x *.sh
./check_env.sh                 # пошаговая диагностика (версия Python, tkinter, …)
./install_deps.sh              # системные пакеты + pip (опционально)
./build.sh                     # сборка → release/out_linux/ProtocolOOT
```

Если комплект лежит на `/mnt/c/...` (WSL + общая папка Windows):

```bash
./sync_workspace.sh            # копия в ~/ProtocolOOT_linux_build
cd ~/ProtocolOOT_linux_build
./check_env.sh && ./build.sh
```

## Быстро (из git, ветка linux)

```bash
git clone -b linux <url>
cd Protocols/linux_port
chmod +x *.sh
./check_env.sh
./build.sh
```

Результат: `release/out_linux/` — `ProtocolOOT`, `data/`, инструкция.

Только бинарник (без комплекта data/):

```bash
source .venv-linux/bin/activate
python release/build_release_linux.py --binary-only --no-verify
```

## Требования

- Python **3.10+** (рекомендуется **3.11–3.12**; 3.14+ для PyInstaller может не подойти)
- `python3-tk`, `binutils` (objdump), `python3-dev` (libpython)

Debian/Ubuntu:

```bash
sudo apt-get install -y python3-venv python3-tk binutils python3-dev
```

ALT Linux (p10):

```bash
sudo apt-get install -y python3.11 python3.11-tools python3-module-pip \
  libpython3.11 python3.11-devel python3-modules-tkinter binutils
rm -rf .venv-linux
./install_deps.sh
./check_env.sh
./build.sh
```

Системный `python3` может быть 3.9 — для сборки используется **python3.11** автоматически.

## Если сборка не стартует

1. `./check_env.sh` — смотрите строки **FAIL** (номер шага и подсказка).
2. `./build.sh --skip-check` — только если check_env уже пройден.
3. `./build.sh --no-verify` — без ruff/verify (если мешают предупреждения).

## Структура комплекта

```text
ProtocolOOT_linux_build/
  app/              ← исходники (готовая Linux-копия)
  release/          ← build_release_linux.py, .spec, out_linux/
  check_env.sh      ← диагностика по шагам
  build.sh          ← сборка
  install_deps.sh
  requirements.txt
  requirements-build.txt
  VERSION.txt
```

Сборку выполняйте в `~/ProtocolOOT_linux_build`, не на смонтированной папке VirtualBox/WSL.
