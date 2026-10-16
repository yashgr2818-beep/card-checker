@echo off
setlocal enabledelayedexpansion
title Business Card Scanner CRM
cd /d "%~dp0"

echo =======================================================
echo         Business Card Scanner CRM - Launcher
echo =======================================================
echo.

:: 1. Check if Python is installed
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Python is not detected on this system!
    echo.
    echo To run this project on this computer:
    echo 1. Download Python from: https://www.python.org/downloads/
    echo 2. During installation, make sure to CHECK the box:
    echo    "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

:: 2. Check if required Python packages are installed
echo [1/3] Checking Python dependencies...
python -c "import flask, PIL, pytesseract, psycopg2, openpyxl, dotenv" >nul 2>nul
if %errorlevel% neq 0 (
    echo [SETUP] First-time setup detected on this system!
    echo Installing required packages from requirements.txt...
    echo.
    pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo [ERROR] Package installation failed. Please check your internet connection.
        pause
        exit /b 1
    )
    echo [OK] All packages installed successfully!
) else (
    echo [OK] All Python dependencies are ready.
)
echo.

:: 3. Check for Tesseract OCR (offline engine)
echo [2/3] Checking OCR Engines...
set TESS_FOUND=0
if exist "C:\Program Files\Tesseract-OCR\tesseract.exe" set TESS_FOUND=1
if exist "C:\Program Files (x86)\Tesseract-OCR\tesseract.exe" set TESS_FOUND=1
where tesseract >nul 2>nul
if %errorlevel% equ 0 set TESS_FOUND=1

if %TESS_FOUND% equ 1 (
    echo [OK] Tesseract OCR detected (Offline scanning available).
) else (
    echo [NOTICE] Tesseract OCR software is not installed on this system.
    echo   - Cloud AI: You can use free Google Gemini Vision AI in Settings.
    echo   - Offline OCR: Install Tesseract from https://github.com/UB-Mannheim/tesseract/wiki
)
echo.

:: 4. Start the application
echo [3/3] Launching web application...
echo Server starting at http://127.0.0.1:5000
echo Your default web browser will open automatically.
echo (Press Ctrl+C to stop the server anytime)
echo.
python app.py

pause
