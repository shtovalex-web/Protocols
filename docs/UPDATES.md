# Автообновление ProtocolOOT (этап 1)

Проверка и установка обновлений для **собранного Windows .exe** по образцу Grafik-PZ.

## Как это работает

1. При старте `.exe` программа **сканирует каталог шары** (`update_config.json` → путь к каталогу).
2. Ищет релизы в **`windows/<версия>/`** и **`linux/<версия>/`** (`manifest.json` и исполняемый файл).
3. Выбирается **новейшая** версия, которая выше текущей (`APP_VERSION`).
4. Если версия новее — показывается диалог установки.
5. Новый `.exe` копируется рядом с текущим как `ProtocolOOT.exe.new`, проверяется **размер** и **SHA-256**.
6. Текущий файл переименовывается в `.old`, новый — в рабочее имя. Программа **не перезапускается сама** — показывается сообщение об успехе; пользователь закрывает окно и снова запускает `ProtocolOOT.exe`. Список изменений — в `data/.pending_changelog.json`, показ после открытия главного окна.
7. Ошибки сети или недоступность шары **не блокируют** запуск (тихий старт).

Даже если в корне шары остался старый `manifest.json`, программа найдёт релизы в `windows/<версия>/manifest.json`.

Ручная проверка: **«Справка» → «Проверить обновления…»**.

## Пути по умолчанию

| Что | Значение |
|-----|----------|
| Каталог шары | `\\SERVER\SOFT\ProtocolOOT` |
| Локальный конфиг | `update_config.json` рядом с `.exe` / в корне проекта |
| Пример манифеста | `docs/update_manifest.example.json` |

## Локальный конфиг `update_config.json`

```json
{
  "manifest_path": "//SERVER/SOFT/ProtocolOOT",
  "enabled": true
}
```

Поле `manifest_path` — **каталог шары** (рекомендуется) или путь к `manifest.json` (устаревший вариант).

## Переменные окружения

| Переменная | Назначение |
|------------|------------|
| `PROTOCOLOOT_UPDATE_MANIFEST` | Путь к манифесту (приоритет над файлом конфигурации) |
| `PROTOCOLOOT_UPDATE_CHECK=1` | Проверять обновления при запуске **из исходников** (для отладки) |

## Формат `manifest.json`

См. `docs/update_manifest.example.json`:

- `latest_version` — версия на шаре (сравнение `major.minor.patch`; `1.5` = `1.5.0`).
- `windows.relative_path` — путь к `.exe` **относительно каталога манифеста**.
- `windows.sha256`, `windows.size` — контроль целостности после копирования.
- `changes_short` — пункты для диалога «Что нового».
- `mandatory` — при `true` диалог «ОК/Отмена» вместо «Да/Нет».

## Публикация обновления на шару

После сборки `.exe`:

```text
py -3 tools/publish_update_manifest.py ^
  --exe "D:\ProtocolOHT_onefile\ProtocolOOT.exe" ^
  --version 1.5.2 ^
  --share-root "\\SERVER\SOFT\ProtocolOOT" ^
  --change "Краткое описание изменения"
```

Скрипт:

1. Кладёт exe и `data/` в `share-root/windows/<версия>/`.
2. Считает SHA-256 и размер.
3. Записывает **`manifest.json` в ту же папку версии** (`windows/<версия>/manifest.json`).

Структура шары:

```text
\\SERVER\SOFT\ProtocolOOT\
  windows\
    1.5.2\
      manifest.json
      ProtocolOOT.exe
      data\
        default_protocol.docx
        FAQ.txt
        …
    1.5.3\
      …
```

## Откат

- **До автообновления:** тег **`v1.5.1`**
- **Стабильная версия с автообновлением (exe + data/):** тег **`v1.5.2`**

## Локальная тестовая шара (Windows)

Для проверки без сетевого UNC можно использовать папку **`D:\Обновление`**:

```text
D:\Обновление\
  windows\
    1.5.2\
      manifest.json
      ProtocolOOT.exe
      data\
```

Рядом с `ProtocolOOT.exe` — `update_config.json`:

```json
{
  "manifest_path": "D:/Обновление",
  "enabled": true
}
```

Проверка манифеста и sha256 без GUI:

```text
py -3 tools/test_update_share.py
py -3 tools/scan_update_share.py
```

`scan_update_share.py` — список всех версий на шаре и рекомендуемое обновление.

При **`py -3 build_windows_exe.py`** (или **`build_windows_exe.bat`**) после сборки:

1. В **`data/`** записывается **`update_info.json`** — маркер версии комплекта (см. `docs/update_info.example.json`).
2. Комлект обновления публикуется на **`D:\Обновление`** в **`windows/<версия>/…`** (Linux — `linux/<версия>/…` позже).
3. Для `ProtocolOHT_onefile` при первой сборке **`update_config.json`** указывает на **`D:/Обновление`**.

Локальная папка **`UPDATE/`** в репозитории **не используется**.

Публикация тестового обновления вручную:

```text
py -3 tools/publish_update_manifest.py ^
  --exe "ProtocolOHT_onefile\ProtocolOOT.exe" ^
  --version 1.5.2 ^
  --share-root "D:\Обновление" ^
  --change "Тестовое обновление"
```

## Предупреждение PyInstaller «Failed to remove temporary directory _MEI…»

При **onefile**-сборке PyInstaller распаковывает exe во временную папку `%TEMP%\_MEI…`.  
Не используйте символ **`!`** в пути к папке с программой — иначе перезапуск после обновления может падать.

С версии **1.6.1** на шаре снова один **`ProtocolOOT.exe`** (не zip). Клиент 1.6.0 с zip-manifest обновится по старой схеме; новые сборки — только exe.

## Этап 2 — обновление `data/`

Вместе с exe обновляются файлы из **`data/`** (см. `update_bundle_files.py`):

- `default_protocol.docx`, `default_protocol_tehnicheskiy.docx`
- инструкции `.docx`, `FAQ.txt`, `ЖУРНАЛ_ДОРАБОТОК.md`
- шаблоны Минтруд XSD (`.xlsx`), `icon.ico`

**Не обновляются:** `protocols.db`, Excel в **корне**, `Protokol/`, `Mintrud/`, `update_config.json`.

**`data/update_info.json`** — локальный маркер версии комплекта; пишется при сборке и после обновления, **не входит** в `data_files` на шаре.

Перед заменой каждого файла резервная копия хранится во **временной папке** (не в `data/`). Устаревшие **`*.bak`** в `data/` удаляются при обновлении. Папка **`data.backup/`** не создаётся. В `manifest.json` — секция **`data_files`**.

После успешного обновления перезаписывается **`data/update_info.json`** (версия комплекта `data/`).

## Ограничения

- Только **Windows .exe** + **`data/`**; Linux — позже.
