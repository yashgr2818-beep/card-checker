@echo off
title Card Scanner CRM (Standalone EXE)
cd /d "%~dp0"

if not exist "dist\CardScanner\CardScanner.exe" (
    echo [ERROR] CardScanner.exe not found in dist\CardScanner\
    echo Please run build_exe.bat first to build the executable.
    pause
    exit /b 1
)

cd dist\CardScanner
start CardScanner.exe
