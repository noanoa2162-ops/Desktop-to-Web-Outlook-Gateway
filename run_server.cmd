@echo off
cd /d "%~dp0"
title Outlook Draft Gateway

set LOG_FILE=%~dp0server_console.log

echo ================================================
echo  Outlook Draft Gateway
echo ================================================
echo.
echo Starting the local web server...
echo Keep this window open while using the app.
echo Log file: %LOG_FILE%
echo.

python server.py 2>&1 | powershell -NoProfile -Command "$input | Tee-Object -FilePath '%LOG_FILE%'"

echo.
echo The server stopped.
echo If there is an error above, send a screenshot of this window.
pause
