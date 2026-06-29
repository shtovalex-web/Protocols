# Пакет desktop_updater

Переиспользуемый модуль автообновления для **собранных Windows .exe** (PyInstaller onefile): проверка `manifest.json` на сетевой шаре, замена exe и файлов `data/`.

ProtocolOOT подключает пакет через `ProtocolOHT_next/protocol_updater_config.py` (вызов `configure(UpdaterConfig(...))`). Старые имена модулей (`update_config`, `startup_update`, …) — тонкие обёртки для обратной совместимости.

## Подключение в другом проекте

1. Скопируйте каталог **`desktop_updater/`** в репозиторий (или подключите submodule).

2. Создайте файл конфигурации (пример):

```python
from pathlib import Path
from desktop_updater import UpdaterConfig, configure

configure(
    UpdaterConfig(
        app_name="MyApp",
        exe_name="MyApp.exe",
        app_version="1.0.0",
        default_share_root=Path(r"\\SERVER\SOFT\MyApp"),
        env_prefix="MYAPP",
        data_subdir="data",
        data_replace_filenames=("config.json", "templates.docx"),
        user_dir_resolver=lambda: Path(__file__).resolve().parent,
    )
)
```

3. **До UI** в `main.py`:

```python
from desktop_updater.startup import prepare_startup_updates
import sys

if not prepare_startup_updates(sys.argv):
    sys.exit(0)
```

4. Рядом с exe — **`update_config.json`** (см. `docs/UPDATES.md`).

5. Публикация релиза — `tools/publish_update_manifest.py` (адаптировать `UpdaterConfig` и список `data_replace_filenames`).

## UpdaterConfig

| Поле | Назначение |
|------|------------|
| `app_name` | Имя для сообщений |
| `exe_name` | Имя exe на шаре (`MyApp.exe`) |
| `app_version` | Версия из исходников (если нет `data/update_info.json`) |
| `default_share_root` | UNC/локальный каталог шары по умолчанию |
| `env_prefix` | Префикс переменных `{PREFIX}_UPDATE_MANIFEST`, `{PREFIX}_UPDATE_CHECK` |
| `data_subdir` | Подпапка комплекта (`data`) |
| `data_replace_filenames` | Файлы для `data_files` в manifest |
| `user_dir_resolver` | Каталог для `update_config.json` (обычно рядом с exe) |
| `restart_cmd_name`, `app_bundle_zip_name` | Legacy onedir/zip (опционально) |

## Структура пакета

```text
desktop_updater/
  config.py          UpdaterConfig
  registry.py        configure(), get_config()
  client_config.py   update_config.json
  manifest.py        manifest.json
  scan.py            windows/<версия>/
  installer.py       копия exe, swap
  data_installer.py  data/
  info.py            update_info.json
  startup.py         prepare_startup_updates()
  …
```

Подробности сценария обновления — **`docs/UPDATES.md`**.
