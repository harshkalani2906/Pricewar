@echo off
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
if not exist "2_Output\_Logs" mkdir "2_Output\_Logs"
python "_engine\qc_price_tracker.py" %* > "2_Output\_Logs\last_scheduled_run.txt" 2>&1
