# Структура репозитория

## Исходники и комплект

| Путь | Назначение |
|------|------------|
| `main.py` | Точка входа |
| `ProtocolOHT_next/` | UI, DOCX, журнал, пути |
| `*.py` в корне | Модули данных (Excel, Минтруд, комиссия…) |
| `bundle/` | **Канонический комплект** шаблонов Word, FAQ, XSD, инструкции |
| `tools/` | Проверка, сборка эталона, синхронизация GitHub, уборка |
| `tests/` | Unit-тесты (`python -m unittest discover -s tests`) |
| `docs/` | Документация для разработчиков (`LINUX_BRANCH.md`, `LINUX_DIST.md`) |

## Портируемые / сборочные копии

| Путь | Назначение |
|------|------------|
| `эталон_сборки/` | Портативная копия для Windows (`tools/update_etalon.py`) |
| `linux_port/` | Порт на Linux; ветка **`linux`** — см. `docs/LINUX_BRANCH.md` |
| `ProtocolOHT_onefile/` | Результат PyInstaller (в `.gitignore`) |

## Рабочие данные (локально, в `.gitignore`)

| Путь | Назначение |
|------|------------|
| `protocols.db`, `last_protocol_no.json` | Журнал и настройки |
| `Data_base.xlsx`, `Programs_base.xlsx` | Базы пользователя |
| `Protokol/`, `Mintrud/` | Сохранённые протоколы и выгрузки |
| `local/` | Черновики |

## Уборка

```bash
python tools/tidy_workspace.py          # что будет удалено
python tools/tidy_workspace.py --apply  # кэши, сборки, linux_port/app, дубликаты bundle/ в корне
python tools/tidy_workspace.py --etalon-only --apply  # мусор в корне эталон_сборки/
```

Или `tidy_workspace.bat` в корне.

Удаляются: `__pycache__`, `.ruff_cache`, `ProtocolOOT_linux_build/`, `ProtocolOOT_linux_build_test_tmp/`, сгенерированный `linux_port/app/`, дубликаты шаблонов в корне (если есть в `bundle/`), случайные `protocols.db` и копии шаблонов в корне `эталон_сборки/` (комплект — в `data/`).

**Не удаляются:** `protocols.db`, Excel и папки `Protokol/`, `Mintrud/` в корне проекта.

## Проверка

```bash
python tools/verify_project.py --no-launch
```

ruff → импорты → unit-тесты.
