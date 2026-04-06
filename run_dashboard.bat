@echo off
title US Economic Tracker
cd /d "%~dp0"

echo.
echo   US Economic Tracker - Web Dashboard
echo   =====================================
echo   Starting Flask server...
echo   Your browser will open automatically.
echo   Close this window to stop the server.
echo.

start "" cmd /c "timeout /t 2 /nobreak >/dev/null && start http://localhost:5000"

py -3 web_dashboard.py

pause >nul
