@echo off
cd /d "%~dp0"
title Outlook Draft Gateway

echo ================================================
echo  Outlook Draft Gateway
echo ================================================
echo.
echo Starting the local web server...
echo Keep this window open while using the app.
echo.

python server.py

echo.
echo The server stopped.
echo If there is an error above, send a screenshot of this window.
pause
