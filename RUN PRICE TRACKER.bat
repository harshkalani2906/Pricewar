@echo off
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
title QC Price Tracker
python "_engine\qc_price_tracker.py" %*
if errorlevel 9009 (
  echo.
  echo  Python is not installed or not on PATH. Install Python 3.10+ from python.org
  echo  ^(tick "Add python.exe to PATH"^), then run FIRST TIME SETUP.bat
)
echo.
pause
