@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Сборка ProtocolOOT.exe (onefile).
echo Перед PyInstaller: ruff check . (нужен ruff: pip install -r requirements-build.txt).
echo Сначала откроется окно выбора папки для exe и шаблонов.
echo Без диалога: перетащите папку на этот bat или укажите путь аргументом.
echo.
py -3 "%~dp0build_windows_exe.py" %*
if errorlevel 1 (
  echo.
  echo Повтор с "python" (если нет команды py):
  python "%~dp0build_windows_exe.py" %*
)
echo.
pause
