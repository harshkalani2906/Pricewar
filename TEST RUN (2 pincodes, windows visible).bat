@echo off
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
title QC Price Tracker - TEST
echo  Test run: first 2 pincodes only, browser windows visible so you can watch.
python "_engine\qc_price_tracker.py" --pins 2 --show
echo.
pause
