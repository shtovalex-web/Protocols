# Linux-дистрибутив (PyInstaller, без исходников)

Готовая папка для пользователя Linux: один файл **`ProtocolOOT`** + каталог **`data/`** (шаблоны, FAQ, Excel-образцы).

## Вариант 1: скачать готовую сборку (GitHub Actions)

1. Откройте репозиторий на GitHub → **Actions** → workflow **build-linux-dist**.
2. Выберите последний успешный запуск (ветка **`linux`**, push или **Run workflow**).
3. Внизу страницы — артефакт **`ProtocolOOT_linux`** (zip).
4. Распакуйте на Linux-ПК, например в `/opt/ProtocolOOT/`.

Структура после распаковки:

```text
ProtocolOOT              ← запускать этот файл
data/                    ← шаблоны, инструкции (не удалять)
Data_base.xlsx           ← образец (если был в сборке)
Programs_base.xlsx
ИНСТРУКЦИЯ_папки_сборки.txt
```

Запуск:

```bash
chmod +x ProtocolOOT
./ProtocolOOT
```

## Вариант 2: собрать самому на Linux

### 2a. Автономный комплект (без git, ~десятки МБ)

На Windows:

```bash
python linux_port/prepare.py
python tools/pack_linux_build.py
```

Скопируйте папку **`ProtocolOOT_linux_build/`** на Linux (zip/USB/сеть).

На Linux:

```bash
cd ProtocolOOT_linux_build
chmod +x *.sh
./install_deps.sh    # системные пакеты и pip (tkinter, libpython, ruff…)
./check_env.sh       # проверка окружения
./build.sh           # или ./build.sh --no-verify при проблемах с ruff
```

Если комплект на `/mnt/c/...` в WSL: `./sync_workspace.sh`, затем сборка в `~/ProtocolOOT_linux_build`.

### 2b. Из git (ветка linux)

```bash
git clone -b linux https://github.com/shtovalex-web/Protocols.git
cd Protocols/linux_port
chmod +x *.sh
./install_deps.sh
./check_env.sh
./build.sh
```

`install_deps.sh` ставит системные пакеты (`python3-tk`, `binutils`, `python3-dev`, LibreOffice, шрифты) и pip-зависимости; `build_linux.sh` — проверки как в [grafik-pz](https://github.com/shtovalex-web/grafik-pz) и запуск PyInstaller.

Результат: `linux_port/release/out_linux/`.

Архив для переноса:

```bash
cd linux_port/release/out_linux
zip -r ../out_linux.zip .
```

## Требования для сборки (Linux)

| Пакет | Зачем |
|-------|--------|
| `python3-tk` | tkinter (проверка перед сборкой) |
| `binutils` | `objdump` для PyInstaller |
| `python3-dev` | `libpython3*.so` (на ALT: **`python3.11-dev`** + **`libpython3.11`**, не `python3.11-devel`) |
| `requirements-build.txt` | PyInstaller, ruff и зависимости приложения (`-r requirements.txt`) |

Debian/Ubuntu: `sudo apt install python3-tk python3-venv python3-dev binutils`

**ALT Linux / p10:** `sudo apt-get install -y python3.11 python3.11-dev libpython3.11 python3-tk binutils`. Для headless-проверок tkinter: **`xorg-xvfb`** (не пакет `xvfb`). LibreOffice: `libreoffice`. Подробнее — `linux_port/release/README_BUILD_LINUX.txt`.

### Если сборка не стартует

| Сообщение | Что сделать |
|-----------|-------------|
| `Не найден …/release/build_release_linux.py` | `git pull` на ветке **linux** (папка `linux_port/release/` должна быть в репозитории) |
| `Нет linux_port/app/main.py` | `python3 linux_port/prepare.py` или клон `-b linux` |
| `No module named openpyxl` / `docx` | `./install_deps.sh` или `pip install -r linux_port/requirements-build.txt` |
| `/usr/bin/env: bash\r` | `python3 fix_crlf.py && chmod +x *.sh` или `sed -i 's/\r$//' *.sh` |
| ruff / verify_linux при сборке | Свежий комплект `ProtocolOOT_linux_build` (в корне есть `ruff.toml`); иначе `./build.sh --no-verify` |
| Ошибки venv/PyInstaller на VirtualBox | Клонируйте проект в `~/Protocols` внутри VM, не на общей папке Windows |

Сборка без проверок: `./build.sh --no-verify` или `./build_linux.sh --no-verify` (в git-клоне).

На ALT Linux и в VM: собирайте в локальной копии (`~/Protocols`), не на смонтированной папке VirtualBox — иначе ломаются симлинки venv/PyInstaller (см. опыт [grafik-pz](https://github.com/shtovalex-web/grafik-pz/blob/main/docs/linux.md)).

## Требования на целевом ПК

| Компонент | Зачем |
|-----------|--------|
| **glibc** той же эпохи, что у сборки (обычно Ubuntu 22.04+ / современный дистрибутив) | совместимость бинарника |
| **python3-tk** не нужен | tkinter вшит в сборку |
| **LibreOffice** (`libreoffice-writer`) | PDF с оформлением DOCX |
| **fonts-dejavu-core**, **fonts-liberation** | кириллица в упрощённом PDF |

Установка на Debian/Ubuntu:

```bash
sudo apt install libreoffice-writer fonts-dejavu-core fonts-liberation
```

## Отличие от запуска из исходников

| | Исходники (`./run.sh`) | Бинарник (`ProtocolOOT`) |
|--|------------------------|---------------------------|
| Нужен git / Python на ПК | да | нет |
| Обновление | `git pull` + `./install_deps.sh` | новый zip из Actions или пересборка |
| Размер | меньше | больше (onefile) |
| Функции | те же | те же (оверлеи Linux уже внутри) |

## Автосборка

Workflow **`.github/workflows/build-linux-dist.yml`** запускается при каждом **push в ветку `linux`** и по кнопке **Run workflow**. После синхронизации `main` → `linux` (см. `docs/LINUX_BRANCH.md`) новый бинарник появится в Actions через несколько минут.
