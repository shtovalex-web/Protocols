@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Снимки экрана (окна программы)...
py -3 "%~dp0tools\capture_manual_screenshots.py"
if errorlevel 1 (
  echo Внимание: снимки не получены — Word будет без картинок или со старыми.
)
echo.
echo Сборка ПОДРОБНАЯ_ИНСТРУКЦИЯ_для_пользователя.docx из .md ...
py -3 "%~dp0tools\generate_podrobnaya_instruction_docx.py"
if errorlevel 1 pause
