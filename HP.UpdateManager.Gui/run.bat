@echo off
set "SCRIPT_DIR=%~dp0"
powershell -ExecutionPolicy Bypass -File "%SCRIPT_DIR%HP.UpdateManager.ps1"
pause
