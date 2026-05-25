@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Установка пакетов из requirements.txt ...
py -3 -m pip install -r requirements.txt
echo.
echo Для PDF через Word: Microsoft Word и py -3 -m pip install pywin32
pause
