@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Обновление папки "эталон_сборки"...
py -3 "%~dp0tools\update_etalon.py" %*
if errorlevel 1 (
  echo.
  echo Повтор с "python" (если нет команды py):
  python "%~dp0tools\update_etalon.py" %*
)
echo.
pause
