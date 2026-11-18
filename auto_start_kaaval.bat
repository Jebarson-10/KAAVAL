@echo off
title KAAVAL - Auto Update Launcher
echo ==================================================
echo   KAAVAL Auto-Updater
echo   Checking GitHub for latest updates...
echo ==================================================
echo.

cd /d "%~dp0"

:: Pull latest changes from GitHub
git pull origin main

echo.
echo ==================================================
echo   Starting KAAVAL...
echo ==================================================

:: Activate virtual environment if it exists
if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
)

:: Run the application
python kaaval_app.py

:: If python crashes or exits, pause so the user can see the error
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo KAAVAL exited with an error code.
    pause
)
