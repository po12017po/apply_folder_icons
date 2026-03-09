@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0apply_folder_icons.ps1"
echo.
echo Done.
pause
