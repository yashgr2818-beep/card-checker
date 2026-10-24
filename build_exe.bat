@echo off
title Build Standalone Executable
cd /d "%~dp0"

echo =======================================================
echo     Building Standalone CardScanner.exe
echo =======================================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python is required on this build machine to compile the executable.
    pause
    exit /b 1
)

echo [1/3] Ensuring PyInstaller is installed...
python -m pip install pyinstaller

echo.
echo [2/3] Compiling app.py into standalone CardScanner.exe...
python -m PyInstaller --noconfirm --onedir --name "CardScanner" app.py

if errorlevel 1 (
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

echo.
echo [3/3] Setting up distribution files...
if exist .env copy /y .env "dist\CardScanner\.env" >nul
if not exist "dist\CardScanner\uploads" mkdir "dist\CardScanner\uploads"

:: Clean up intermediate compile artifacts to save disk space
if exist build rmdir /s /q build
if exist CardScanner.spec del /q CardScanner.spec

echo.
echo =======================================================
echo [SUCCESS] Standalone app created in: dist\CardScanner\
echo You can run "dist\CardScanner\CardScanner.exe" directly!
echo No Python installation is needed on target computers.
echo =======================================================
echo.
pause
