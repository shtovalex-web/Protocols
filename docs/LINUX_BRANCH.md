# Ветка linux: полная копия приложения

Ветка **`linux`** — та же кодовая база, что и **`main`**, плюс **зафиксированная** готовая копия в `linux_port/app/` (на `main` она только генерируется и в git не попадает).

## Схема

```text
main (Windows) ──merge──► linux
       │                    │
       │  prepare.py        │  linux_port/app/  (полная копия + оверлеи)
       └────────────────────┘
```

Оверлеи Linux (PDF через LibreOffice, шрифты) — только в `linux_port/overlays/`, накладываются при `prepare.py`.

## Ручная синхронизация

После правок на `main`:

```bash
python tools/sync_linux_branch.py
python tools/sync_linux_branch.py --push
```

Или `sync_linux_branch.bat`.

Скрипт: merge `main` → `linux`, `prepare.py`, проверка, коммит `linux_port/app/`.

Незакоммиченные правки на диске: `python tools/sync_linux_branch.py --allow-dirty` (prepare читает файлы с диска).

## Автоматически (GitHub)

При **push в `main`** workflow `.github/workflows/sync-linux.yml`:

1. переключается на ветку `linux`;
2. вливает `main`;
3. запускает `prepare.py`;
4. коммитит и пушит `linux`, если есть изменения.

## Развёртывание на Linux

```bash
git clone -b linux https://github.com/shtovalex-web/Protocols.git
cd Protocols/linux_port
chmod +x *.sh
./install_deps.sh
./run.sh
```

Сборка бинарника: `./build_linux.sh` (только на Linux).

## Проверка

```bash
python linux_port/prepare.py
python linux_port/verify_linux.py --no-launch
python -m unittest discover -s tests -v
```

На CI ветки `linux` дополнительно гоняется job на `ubuntu-latest`.
