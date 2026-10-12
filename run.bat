@echo off
setlocal enabledelayedexpansion
title Local Word ^<--^> PDF Converter

echo =========================================================
echo       Local Word ^<--^> PDF Converter for Windows
echo                   100%% Offline ^& Secure
echo =========================================================
echo.

:: 1. Verify Python installation
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not found in your system PATH.
    echo.
    echo Please install Python 3.10 or newer from:
    echo https://www.python.org/downloads/
    echo.
    echo IMPORTANT: Make sure to check "Add python.exe to PATH" during installation.
    echo.
    pause
    exit /b 1
)

:: 2. Check if dependencies are installed
echo Checking required dependencies...
python -c "import win32com, pdf2docx" >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo Dependencies appear to be missing. Installing now from requirements.txt...
    python -m pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo.
        echo [ERROR] Failed to install dependencies. Please run 'pip install -r requirements.txt' manually.
        pause
        exit /b 1
    )
    echo Dependencies successfully installed!
    echo.
)

:: 3. Launch the application
echo Starting application...
start "" pythonw main.py
if %errorlevel% neq 0 (
    :: Fallback to python if pythonw is not available
    python main.py
)

exit /b 0
