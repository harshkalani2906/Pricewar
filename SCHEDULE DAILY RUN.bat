@echo off
cd /d "%~dp0"
title Price War - schedule automatic runs
if not exist "2_Output\_Logs" mkdir "2_Output\_Logs"
echo.
echo  Windows will now run Price War by itself:
echo    - every day at 07:30          : full check (80%% revenue SKUs + all QC apps)
echo    - every 2 hours, day and night : Top 50 SKUs + all QC apps
echo  After every run the dashboard on GitHub updates by itself.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0_engine\schedule.ps1"
if errorlevel 1 goto fail
echo.
echo  Laptop settings so it keeps working with the LID CLOSED:
echo    - closing the lid does nothing (no sleep), on charger and on battery
echo    - no sleep while on the charger
echo    - wake-up timers ON, so a sleeping laptop still wakes up for each run
powercfg /setacvalueindex SCHEME_CURRENT SUB_BUTTONS LIDACTION 0
powercfg /setdcvalueindex SCHEME_CURRENT SUB_BUTTONS LIDACTION 0
powercfg /setacvalueindex SCHEME_CURRENT SUB_SLEEP RTCWAKE 1
powercfg /setdcvalueindex SCHEME_CURRENT SUB_SLEEP RTCWAKE 1
powercfg /change standby-timeout-ac 0
powercfg /setactive SCHEME_CURRENT
echo.
echo  ==== DONE ====
echo  Keep the laptop ON and on the CHARGER, and stay logged in to Windows
echo  (a locked screen is fine, closing the lid is fine - do not Shut down or Sign out).
echo  To stop: double-click "STOP DAILY RUN.bat"
pause
exit /b
:fail
echo.
echo  Could not create the schedule. Right-click this file - "Run as administrator" - and try again.
echo  Still failing? Send a screenshot of this window to Claude.
pause
