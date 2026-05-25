# Синхронизация с GitHub

## Быстрый старт

1. Убедитесь, что настроен `origin` (уже: `shtovalex-web/Protocols`).
2. Сохраните изменения в файлах.
3. Запустите **`sync_github.bat`** или:
   ```bash
   py -3 tools/sync_github.py -m "что изменили"
   ```

## Авто-push после commit

Один раз выполните **`setup_git_hooks.bat`** — в локальном репозитории будет `core.hooksPath = .githooks`.

После каждого `git commit` хук `post-commit` выполнит `git push origin <текущая_ветка>`.

## Что не попадает в репозиторий

См. `.gitignore`: рабочие базы (`Data_base.xlsx`, `protocols.db`), папка `Protokol/`, сборка `ProtocolOHT_onefile/`, пакеты `ib_*.zip`, кэши, `.cursor/`.

## CI на GitHub

Workflow `.github/workflows/verify.yml` — ruff и импорты при push в `main`.
