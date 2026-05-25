@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Синхронизация с GitHub...
py -3 "%~dp0tools\sync_github.py" %*
if errorlevel 1 (
  echo.
  echo Повтор с "python":
  python "%~dp0tools\sync_github.py" %*
)
echo.
pause
