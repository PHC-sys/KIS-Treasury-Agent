@echo off
chcp 65001 >nul
cd /d "%~dp0"
python src\run.py
echo.
echo Done. Press any key to close.
pause >nul
