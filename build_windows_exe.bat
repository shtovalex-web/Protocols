@echo off
chcp 65001 >nul
cd /d "%~dp0"

rem Куда складывать сборку: DEPLOY_ROOT\<версия>\ (версия — APP_VERSION в protocol_app_info.py)
set "DEPLOY_ROOT=D:\Проекты Курсор\Протоколы по ОТ\Протокола по ОТ 1.5\Windows"

echo Сборка ProtocolOOT.exe (onefile).
echo Перед PyInstaller: ruff check . (нужен ruff: pip install -r requirements-build.txt).
echo После сборки: update_info.json в data\ и публикация в D:\Обновление\windows\^<версия^>\.
echo.

set "BUILD_ARGS=%*"
if not "%~1"=="" goto run_build

for /f "delims=" %%V in ('py -3.12 "%~dp0build_windows_exe.py" --print-version 2^>nul') do set "APP_VERSION=%%V"
if not defined APP_VERSION for /f "delims=" %%V in ('py -3 "%~dp0build_windows_exe.py" --print-version 2^>nul') do set "APP_VERSION=%%V"
if not defined APP_VERSION for /f "delims=" %%V in ('python "%~dp0build_windows_exe.py" --print-version 2^>nul') do set "APP_VERSION=%%V"
if not defined APP_VERSION (
  echo Ошибка: не удалось прочитать APP_VERSION из ProtocolOHT_next\protocol_app_info.py
  pause
  exit /b 1
)

set "BUILD_OUT=%DEPLOY_ROOT%\%APP_VERSION%"
echo Папка сборки: %BUILD_OUT%
set "BUILD_ARGS=%BUILD_OUT%"

:run_build
py -3.12 "%~dp0build_windows_exe.py" %BUILD_ARGS%
if errorlevel 1 (
  echo.
  echo Повтор с py -3:
  py -3 "%~dp0build_windows_exe.py" %BUILD_ARGS%
)
if errorlevel 1 (
  echo.
  echo Повтор с "python":
  python "%~dp0build_windows_exe.py" %BUILD_ARGS%
)
echo.
pause
