@echo off
title Genealogy Viewer - Build EXE
cd /d "%~dp0"

set "PY=C:\Users\Administrator\AppData\Local\Programs\Python\Python312\python.exe"
if not exist "%PY%" (
    echo [!] Python 3.12 not found, fallback to python in PATH
    set "PY=python"
)

echo ================================================
echo    Genealogy Viewer - One-click Build
echo ================================================
echo   Python: %PY%
echo.

echo [1/3] Checking PyInstaller ...
"%PY%" -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo   Not installed. Installing PyInstaller ...
    "%PY%" -m pip install pyinstaller
    if errorlevel 1 (
        echo   [!] PyInstaller install failed. Check network and run again.
        pause
        exit /b 1
    )
) else (
    echo   Installed.
)

echo.
echo [2/3] Building ... takes 1-2 minutes, do NOT close this window.
"%PY%" build_exe.py
if errorlevel 1 (
    echo.
    echo [!] Build FAILED. Copy the error messages above for feedback.
    pause
    exit /b 1
)

echo.
echo ================================================
echo   Build OK!
echo   EXE location: %~dp0dist_new\
echo.
echo   Note: put the EXE together with your genealogy .json
echo   file, or double-click the json and choose this EXE.
echo ================================================
pause
