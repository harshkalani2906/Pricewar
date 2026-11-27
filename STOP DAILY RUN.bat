@echo off
schtasks /delete /f /tn "Price War - daily 80pc" >nul 2>&1
schtasks /delete /f /tn "Price War - Top 50 every 2 hours" >nul 2>&1
schtasks /delete /f /tn "Price War - Top 50 every 40 min" >nul 2>&1
schtasks /delete /f /tn "QC Price Tracker daily" >nul 2>&1
echo  Automatic runs stopped.
pause
