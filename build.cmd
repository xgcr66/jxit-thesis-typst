@echo off
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0build.ps1" -UpdateFields
if errorlevel 1 (echo Build failed. See details above.) else (start "" "%~dp0build")
pause
