# Тесты

```bash
python -m unittest discover -s tests -v
```

| Файл | Что проверяет |
|------|----------------|
| `test_mintrud_export_v.py` | Выгрузка Минтруд, программа «В» |
| `test_protocol_output_linux.py` | Linux-порт: PDF и LibreOffice (нужен `linux_port/prepare.py`) |

Проверка проекта целиком: `python tools/verify_project.py --no-launch` (ruff, импорты, тесты).
