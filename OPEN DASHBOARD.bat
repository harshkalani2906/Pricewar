@echo off
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
echo  Updating the dashboard with the latest prices...
python "_engine\build_dashboard.py"
start "" "%~dp03_Dashboard\price-pulse.html"
