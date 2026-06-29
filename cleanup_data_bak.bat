@echo off
chcp 65001 >nul
cd /d "%~dp0"
if "%~1"=="" (
  echo Удаление *.bak из data\ перед обновлением ProtocolOOT.
  echo.
  echo Перетащите папку с ProtocolOOT.exe на этот bat
  echo или: cleanup_data_bak.bat "путь\к\папке"
  pause
  exit /b 1
)
py -3 "%~dp0tools\cleanup_data_bak.py" %*
if errorlevel 1 python "%~dp0tools\cleanup_data_bak.py" %*
echo.
pause
