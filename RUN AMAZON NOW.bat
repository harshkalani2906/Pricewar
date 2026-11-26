@echo off
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
title Price War - Amazon only
echo.
echo  Checking Amazon prices for every ASIN in 1_Input\Links.xlsx ...
echo.
python "_engine\qc_price_tracker.py" --only amazon
start "" "%~dp03_Dashboard\price-pulse.html"
echo.
pause
