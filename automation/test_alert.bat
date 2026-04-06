@echo off
title Alert Test
cd /d "%~dp0"
echo.
py -3 test_alert.py
echo.
pause
