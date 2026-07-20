@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "MODE=local"
if not "%1"=="" set "MODE=%1"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0start.ps1" -Mode "%MODE%"
pause
