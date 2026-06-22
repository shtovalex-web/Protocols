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

Проверка проекта целиком: `python tools/verify_project.py --no-launch` (ruff, импорты, все тесты).
