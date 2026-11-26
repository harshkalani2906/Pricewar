@echo off
cd /d "%~dp0"
title Price War - schedule automatic runs
if not exist "2_Output\_Logs" mkdir "2_Output\_Logs"
echo.
echo  This makes Windows run Price War by itself (the laptop must be ON; a locked screen is fine):
echo.
echo    1. Every day at 07:30  - full check: all apps + marketplaces for the 80% revenue SKUs
echo    2. Every 2 hours, 09:30 to 21:30 - Top 50 revenue SKUs (+ all quick-commerce apps)
echo.
schtasks /delete /f /tn "QC Price Tracker daily" >nul 2>&1
schtasks /create /f /tn "Price War - daily 80pc" /sc daily /st 07:30 /tr "\"%~dp0RUN IN BACKGROUND.bat\" --tier 80"
schtasks /create /f /tn "Price War - Top 50 every 2 hours" /sc daily /st 09:30 /ri 120 /du 12:01 /tr "\"%~dp0RUN IN BACKGROUND.bat\" --tier top50"
echo.
if %errorlevel%==0 (echo  DONE - both schedules are set.) else (echo  Could not create the schedule - send a screenshot to Claude.)
echo  To stop them later: double-click "STOP DAILY RUN.bat"
pause
