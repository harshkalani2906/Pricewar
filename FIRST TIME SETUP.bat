@echo off
cd /d "%~dp0"
title Price Parity Tracker - first time setup
echo.
echo  Installing what the QC Price Tracker needs. Takes a few minutes, once.
echo.
python -m pip install --upgrade pandas openpyxl xlsxwriter requests curl_cffi beautifulsoup4 playwright
python -m playwright install chromium
echo.
echo  ==== SETUP FINISHED ====
echo  Next: fill the 4 Excel files in 1_Input (Links, Product Master, Pincodes, Pincode Master),
echo  then double-click "TEST RUN (2 pincodes, windows visible).bat"
pause
