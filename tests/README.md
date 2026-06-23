# Тесты

```bash
python tools/verify_project.py --no-launch
```

или из каталога `tests/`:

```bash
cd tests
python -m unittest discover -v
```

| Файл | Что проверяет |
|------|----------------|
| `test_journal_upsert.py` | Журнал: upsert, dedupe списка, purge дублей |
| `test_commission_profiles.py` | Профили комиссии в `protocols.db` |
| `test_mintrud_export_merge.py` | Dedupe/merge строк выгрузки Минтруд |
| `test_mintrud_export_v.py` | Выгрузка Минтруд, программа «В» |
| `test_protocol_output_linux.py` | Linux-порт: PDF и LibreOffice (нужен `linux_port/prepare.py`) |
| `test_build_release_linux.py` | Комплект Linux-сборки (`assemble_release`) |
| `test_pack_linux_build.py` | Автономный комплект `ProtocolOOT_linux_build/` |
| `test_sync_linux_local.py` | Какие файлы требуют sync Linux-копии |
| `test_ui_theme.py` | Единая тема tkinter (`ui_theme.py`, `ui_widgets.py`), палитры, кнопки |
| `test_protocol_docx_smoke.py` | Сборка DOCX по шаблону `default_protocol.docx` |

Проверка проекта целиком: `python tools/verify_project.py --no-launch` (ruff, импорты, все тесты).

Тесты с SQLite (`test_commission_profiles`, `test_journal_upsert`) закрывают соединения через `_bootstrap.close_tracked_sqlite_connections()` — иначе на Windows файл БД остаётся заблокированным.
