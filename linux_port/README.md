# Linux-порт программы формирования протоколов

Отдельная ветка **`linux`** и папка `linux_port/` — полная копия для Linux без правок Windows-исходников в `main`.

## Ветки git

| Ветка | Назначение |
|-------|------------|
| `main` | Разработка под Windows; `linux_port/app/` не в git |
| `linux` | То же + **зафиксированная** копия `linux_port/app/` после каждой синхронизации |

Подробно: `docs/LINUX_BRANCH.md`.

**Синхронизация после правок на main:**

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

```bash
./install_deps.sh
./build_linux.sh
```

`install_deps.sh` ставит `python3-tk`, `binutils`, `python3-dev` и pip-зависимости (как в [grafik-pz](https://github.com/shtovalex-web/grafik-pz)).

Результат: **`ProtocolOHT_linux_dist/`** — `ProtocolOOT` + `data/`.

Готовый **zip** с GitHub: Actions → **build-linux-dist** → артефакт `ProtocolOHT_linux_dist.zip`.  
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
