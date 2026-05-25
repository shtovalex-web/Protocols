@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Обновление ИНСТРУКЦИЯ_оформление_протоколов_Минтруд.docx из .md ...
py -3 "%~dp0tools\generate_oformlenie_instruction_docx.py"
if errorlevel 1 pause
