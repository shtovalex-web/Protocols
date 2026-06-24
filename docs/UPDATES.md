# Автообновление ProtocolOOT (этап 1)

Проверка и установка обновлений для **собранного Windows .exe** по образцу Grafik-PZ.

## Как это работает

1. При старте `.exe` программа читает `manifest.json` и **сканирует вложенные каталоги** шары.
2. Ищет `manifest.json` и `ProtocolOOT.exe` в папках с именем версии (`windows/1.5.2/…`).
3. Выбирается **новейшая** версия, которая выше текущей (`APP_VERSION`).
4. Если версия новее — показывается диалог установки.
5. Новый `.exe` копируется рядом с текущим как `ProtocolOOT.exe.new`, проверяется **размер** и **SHA-256**.
6. Текущий файл переименовывается в `.old`, новый — в рабочее имя; запускается обновлённый процесс с `--show-changelog=версия`.
7. Ошибки сети или недоступность шары **не блокируют** запуск (тихий старт).

Даже если корневой `manifest.json` устарел, программа найдёт более новый `ProtocolOOT.exe` во вложенных папках `windows/<версия>/`.

Ручная проверка: **«Справка» → «Проверить обновления…»**.

## Пути по умолчанию

| Что | Значение |
|-----|----------|
| Манифест на шаре | `\\SERVER\SOFT\ProtocolOOT\manifest.json` |
| Локальный конфиг | `update_config.json` рядом с `.exe` / в корне проекта |
| Пример манифеста | `docs/update_manifest.example.json` |

## Локальный конфиг `update_config.json`

```json
{
  "manifest_path": "\\\\SERVER\\SOFT\\ProtocolOOT\\manifest.json",
  "enabled": true
}
```

Поле `manifest_path` — UNC или локальный путь к `manifest.json`.

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

1. Кладёт exe в `share-root/windows/<версия>/ProtocolOOT.exe`.
2. Считает SHA-256 и размер.
3. Записывает `manifest.json` в корень шары.

## Откат

Стабильная версия до автообновления: **тег `v1.5.1`** в репозитории.

## Локальная тестовая шара (Windows)

Для проверки без сетевого UNC можно использовать папку **`D:\Обновление`**:

```text
D:\Обновление\
  manifest.json
  windows\1.5.2\ProtocolOOT.exe
```

Рядом с `ProtocolOOT.exe` — `update_config.json`:

```json
{
  "manifest_path": "D:\\Обновление\\manifest.json",
  "enabled": true
}
```

Проверка манифеста и sha256 без GUI:

```text
py -3 tools/test_update_share.py
py -3 tools/scan_update_share.py
```

`scan_update_share.py` — список всех версий на шаре и рекомендуемое обновление.

Публикация тестового обновления:

```text
py -3 tools/publish_update_manifest.py ^
  --exe "ProtocolOHT_onefile\ProtocolOOT.exe" ^
  --version 1.5.2 ^
  --share-root "D:\Обновление" ^
  --change "Тестовое обновление"
```

## Предупреждение PyInstaller «Failed to remove temporary directory _MEI…»

При **onefile**-сборке PyInstaller распаковывает exe во временную папку `%TEMP%\_MEI…`.  
После **автообновления** (перезапуск exe) Windows иногда показывает предупреждение, что эту папку не удалось стереть. Это **не ошибка программы** — на работу ProtocolOOT не влияет; Windows очистит `%TEMP%` позже.

С версии 1.5.2 перезапуск после обновления идёт через `cmd start` (отдельный процесс), чтобы таких предупреждений было меньше.

## Этап 2 — обновление `data/`

Вместе с exe обновляются файлы из **`data/`** (см. `update_bundle_files.py`):

- `default_protocol.docx`, `default_protocol_tehnicheskiy.docx`
- инструкции `.docx`, `FAQ.txt`, `ЖУРНАЛ_ДОРАБОТОК.md`
- шаблоны Минтруд XSD (`.xlsx`), `icon.ico`

**Не обновляются:** `protocols.db`, Excel в **корне**, `Protokol/`, `Mintrud/`, `update_config.json`.

Перед заменой создаётся **`data.backup/`**. В `manifest.json` — секция **`data_files`**.

## Ограничения

- Только **Windows .exe** + **`data/`**; Linux — позже.
