@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Настройка git-хуков: push на GitHub после каждого commit...
git config core.hooksPath .githooks
if errorlevel 1 (
  echo Ошибка настройки hooksPath.
  pause
  exit /b 1
)
echo Готово. hooksPath = .githooks
echo После commit будет выполняться: git push origin ^<текущая ветка^>
echo.
pause
