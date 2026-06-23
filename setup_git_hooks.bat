@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Настройка git-хуков: pre-commit (Linux sync) + push после commit...
git config core.hooksPath .githooks
if errorlevel 1 (
  echo Ошибка настройки hooksPath.
  pause
  exit /b 1
)
echo Готово. hooksPath = .githooks
echo Перед commit: sync linux_port/app при правках исходников
echo После commit: git push origin ^<текущая ветка^>
echo.
pause
