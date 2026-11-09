@echo off
chcp 65001 >nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0CAI-DAT-WINDOWS.ps1"
echo.
pause
