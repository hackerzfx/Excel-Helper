@echo off
title Excel Helper
cd /d "%~dp0"

where npm >nul 2>&1
if errorlevel 1 (
    echo Node.js is required to run Excel Helper.
    echo Please install it from https://nodejs.org and then run this file again.
    pause
    exit /b 1
)

if not exist node_modules (
    echo First run: downloading Excel Helper components, please wait a few minutes...
    call npm install --no-audit --no-fund
)

echo Starting Excel Helper... close the window that opens to quit the app.
call npm start
