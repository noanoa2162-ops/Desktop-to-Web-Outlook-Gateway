@echo off
setlocal
cd /d "%~dp0"
title Outlook Draft Server

echo ================================================
echo   Outlook draft local server
echo ================================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo ERROR: Python was not found.
    echo Install Python and run this file again.
    pause
    exit /b 1
)

echo Checking Python packages...
python -c "import flask, flask_cors, win32com.client" >nul 2>nul
if errorlevel 1 (
    echo Missing packages. Installing requirements...
    python -m pip install -r requirements.txt --quiet
    if errorlevel 1 (
        echo ERROR: Could not install Python packages.
        echo Try running this command manually:
        echo python -m pip install -r requirements.txt
        pause
        exit /b 1
    )
)

echo Cleaning old local Python servers...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-NetTCPConnection -LocalAddress 127.0.0.1 -State Listen -ErrorAction SilentlyContinue | Where-Object { $_.LocalPort -ge 5000 -and $_.LocalPort -le 5019 } | ForEach-Object { $p = Get-Process -Id $_.OwningProcess -ErrorAction SilentlyContinue; if ($p -and $p.ProcessName -like 'python*') { Stop-Process -Id $p.Id -Force } }" >nul 2>nul

if not defined OPEN_BROWSER set OPEN_BROWSER=1

echo Starting local server...
echo The browser should open automatically.
echo If it does not, copy the URL printed below.
echo Stop the server with Ctrl+C.
echo.
python server.py

echo.
echo Server stopped.
pause
