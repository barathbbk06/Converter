@echo off
setlocal enabledelayedexpansion
title DocMorph Web App Server

echo =========================================================
echo       DocMorph Web Converter (Local & Cloud Ready)
echo =========================================================
echo.

:: 1. Verify Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not in PATH.
    pause
    exit /b 1
)

:: 2. Launch web server
echo Starting web server at http://127.0.0.1:5000...
echo Open your web browser to http://127.0.0.1:5000
echo.
start http://127.0.0.1:5000
python web_app.py

pause
