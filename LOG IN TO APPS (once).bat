@echo off
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
title Log in to Instamart, BigBasket and Flipkart Minutes (once)
echo.
echo  This opens Instamart, then BigBasket, then Flipkart Minutes - one window at a time.
echo  Log in to each with your phone number and OTP. You only do this once;
echo  the tracker remembers the login after that.
echo.
python "_engine\qc_price_tracker.py" --login instamart
python "_engine\qc_price_tracker.py" --login bigbasket
python "_engine\qc_price_tracker.py" --login fkminutes
echo.
echo  ==== ALL DONE - now run the TEST RUN ====
pause
