@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo  KTB intraday (5min) collector
echo  first run = backfill, then = incremental
echo  (Infomax login required. Close infomax_data.xlsx first.)
echo ============================================================
python src\intraday_pull.py sync
echo.
echo Done. Press any key to close.
pause >nul
