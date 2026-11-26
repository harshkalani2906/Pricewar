@echo off
cd /d "%~dp0"
title Price War - push to GitHub
python "_engine\push_github.py" %*
echo.
pause
