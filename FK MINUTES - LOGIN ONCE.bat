@echo off
cd /d "%~dp0"
echo  A Flipkart window will open. Log in with OTP, then come back here and press Enter.
python "_engine\qc_price_tracker.py" --login fkminutes
pause
