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
| `docs/` | Документация для разработчиков |

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
python tools/tidy_workspace.py --apply  # кэши, сборки, linux_port/app
```

Или `tidy_workspace.bat` в корне.

## Проверка

```bash
python tools/verify_project.py --no-launch
```

ruff → импорты → unit-тесты.
