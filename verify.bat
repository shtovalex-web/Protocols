@echo off
chcp 65001 >nul
cd /d "%~dp0"
python "%~dp0tools\verify_project.py" %*
exit /b %ERRORLEVEL%
