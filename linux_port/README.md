# Linux-порт программы формирования протоколов

Отдельная ветка **`linux`** и папка `linux_port/` — полная копия для Linux без правок Windows-исходников в `main`.

## Ветки git

| Ветка | Назначение |
|-------|------------|
| `main` | Разработка под Windows; `linux_port/app/` не в git |
| `linux` | То же + **зафиксированная** копия `linux_port/app/` после каждой синхронизации |

Подробно: `docs/LINUX_BRANCH.md`.

**Синхронизация после правок на main:**

Локально (автоматически при `python tools/verify_project.py` или вручную):

```bash
python tools/sync_linux_local.py
```

Обновляет `linux_port/app/` и `ProtocolOOT_linux_build/`. Перед `git commit` — хук `pre-commit` (см. `setup_git_hooks.bat`).

В git (ветка `linux`):

```bash
python tools/sync_linux_branch.py
python tools/sync_linux_branch.py --push
```

На GitHub при push в `main` ветка `linux` обновляется автоматически (workflow `sync-linux.yml`).

## Что внутри `linux_port/`

| Файл / каталог | Назначение |
|----------------|------------|
| `prepare.py` | Копирует исходники из корня в `app/` и накладывает Linux-оверлеи |
| `overlays/` | Замены модулей (PDF через LibreOffice, шрифты Linux) |
| `app/` | Полная копия приложения (на ветке `linux` — в git) |
| `install_deps.sh`, `run.sh`, `build_linux.sh` | Установка, запуск, сборка PyInstaller |
| `verify_linux.py` | Проверка копии |

## Клонирование для Linux

```bash
git clone -b linux <url-репозитория>
cd Protocols/linux_port
chmod +x *.sh
./install_deps.sh
./run.sh
```

## Бинарник (дистрибутив без исходников)

### Вариант A: автономный комплект (без всего git)

На Windows перед переносом на Linux:

```bash
python linux_port/prepare.py
python tools/pack_linux_build.py
```

Папка **`ProtocolOOT_linux_build/`** (~десятки МБ) — скопируйте на Linux, не весь репозиторий.

На Linux:

```bash
cd ProtocolOOT_linux_build
python3 fix_crlf.py     # если «bash\r» после копирования с Windows
chmod +x *.sh
./check_env.sh
./install_deps.sh       # при необходимости
./build.sh
```

Если комплект на `/mnt/c/...` (WSL): `./sync_workspace.sh` → `~/ProtocolOOT_linux_build`.

Подробно: `linux_port/release/README_BUILD_LINUX.txt`.

### Вариант B: из git (ветка linux)

```bash
./install_deps.sh
./check_env.sh
./build.sh
```

Результат: **`linux_port/release/out_linux/`** — `ProtocolOOT`, `data/`, инструкция (как [grafik-pz](https://github.com/shtovalex-web/grafik-pz) `release/out_linux/`).

Готовый **zip** с GitHub: Actions → **build-linux-dist** → артефакт `ProtocolOOT_linux`.  
Подробно: `docs/LINUX_DIST.md`.

## Функциональность

| Возможность | Linux-порт |
|-------------|------------|
| Интерфейс tkinter | да |
| DOCX, Excel, SQLite, Минтруд | да |
| PDF с оформлением DOCX | LibreOffice, запасной docx2pdf |
| Упрощённый PDF | системные TTF |
| Word COM / pywin32 | не используется |

Оверлеи только для PDF/шрифтов; остальной код синхронизируется с `main` через `prepare.py`.

## Проверка

```bash
python linux_port/prepare.py
python linux_port/verify_linux.py --no-launch
python -m unittest discover -s tests -v
```
