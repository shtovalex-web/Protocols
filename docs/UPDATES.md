# Автообновление ProtocolOOT (этап 1)

Проверка и установка обновлений для **собранного Windows .exe** по образцу Grafik-PZ.

## Как это работает

1. При старте `.exe` программа читает `manifest.json` с сетевой шары (или путь из конфигурации).
2. Если `latest_version` новее текущей (`APP_VERSION` в `protocol_app_info.py`), показывается диалог установки.
3. Новый `.exe` копируется рядом с текущим как `ProtocolOOT.exe.new`, проверяется **размер** и **SHA-256**.
4. Текущий файл переименовывается в `.old`, новый — в рабочее имя; запускается обновлённый процесс с `--show-changelog=версия`.
5. Ошибки сети или отсутствие манифеста **не блокируют** запуск (тихий старт).

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
  --version 1.6.0 ^
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
  windows\1.6.1\ProtocolOOT.exe
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
```

Публикация тестового обновления:

```text
py -3 tools/publish_update_manifest.py ^
  --exe "ProtocolOHT_onefile\ProtocolOOT.exe" ^
  --version 1.6.1 ^
  --share-root "D:\Обновление" ^
  --change "Тестовое обновление"
```

## Ограничения этапа 1

- Только **Windows .exe** (без zip-пакета и без обновления каталога `data/`).
- Linux и обновление данных пользователя — следующие этапы.
