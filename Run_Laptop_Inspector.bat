@echo off
echo Starting Laptop Inspection...
echo This will take 2-3 minutes to complete.
echo.
pause
powershell -ExecutionPolicy Bypass -File "%~dp0Laptop_Inspector.ps1"
