
@echo off
setlocal
title DLT Statistics Collector

REM Save as START_DLT.bat in the same folder as dlt_collector.py

set "PROJECT_DIR=%~dp0"
set "SCRIPT=%PROJECT_DIR%dlt_collector.py"

REM Check application
if not exist "%SCRIPT%" (
    echo [ERROR] dlt_collector.py not found.
    echo Place START_DLT.bat in the DLT project folder.
    pause
    exit /b 1
)

REM Detect Python
set "PYTHON=python"

if exist "%PROJECT_DIR%..\portable\python\python.exe" (
    set "PYTHON=%PROJECT_DIR%..\portable\python\python.exe"
) else (
    where python >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Python not found.
        pause
        exit /b 1
    )
)

REM Launch Windows Terminal in true fullscreen mode
where wt.exe >nul 2>&1

if not errorlevel 1 (
    wt.exe --window new --fullscreen new-tab --title "DLT Statistics Collector" --startingDirectory "%PROJECT_DIR%" "%PYTHON%" "%SCRIPT%"
    exit /b 0
)

REM Fallback: Maximized PowerShell/Command Prompt window
cd /d "%PROJECT_DIR%"
start "DLT Statistics Collector" /MAX cmd.exe /k ""%PYTHON%" "%SCRIPT%""

exit /b 0
