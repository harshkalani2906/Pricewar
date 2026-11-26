@echo off
schtasks /delete /f /tn "Price War - daily 80pc"
schtasks /delete /f /tn "Price War - Top 50 every 2 hours"
schtasks /delete /f /tn "QC Price Tracker daily" >nul 2>&1
echo  Automatic runs stopped.
pause
