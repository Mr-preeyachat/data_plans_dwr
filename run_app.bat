@echo off
chcp 65001 > NUL
title DWR007 Project Dashboard Server
echo ===================================================
echo   DWR007 Project Dashboard - Starting Server...
echo ===================================================
echo.
python "%~dp0server.py"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Failed to start server. Please ensure Python is installed.
    pause
)
