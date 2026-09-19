@echo off
chcp 65001 >nul
cd /d "%~dp0"
python src\run.py
echo.
echo ============================================================
echo  5min bars (intraday) - KTB futures(C65/C67) + UST(2Y/10Y)
echo ============================================================
python src\intraday_pull.py sync C65 C67
python src\intraday5_ust.py sync
python src\intraday5_fx.py sync
python src\push_intraday5.py
echo.
echo Done. Press any key to close.
pause >nul
