@echo off
title Laptop Inspection Tool
color 0A
echo.
echo  ╔══════════════════════════════════════════════╗
echo  ║           LAPTOP INSPECTION TOOL             ║
echo  ║                                              ║
echo  ║  This will test the laptop hardware and      ║
echo  ║  generate a detailed HTML report.            ║
echo  ║                                              ║
echo  ║  Time required: 2-3 minutes                  ║
echo  ║  Administrator rights recommended            ║
echo  ╚══════════════════════════════════════════════╝
echo.
echo Press any key to start the inspection...
pause >nul

echo.
echo Starting PowerShell script...
echo.

REM Run PowerShell with proper execution policy
powershell.exe -ExecutionPolicy Bypass -WindowStyle Normal -File "%~dp0Laptop_Inspector_Fixed.ps1"

REM Check if PowerShell failed
if %ERRORLEVEL% neq 0 (
    echo.
    echo ❌ PowerShell script encountered errors.
    echo.
    echo Trying alternative method...
    echo.
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& {Set-Location '%~dp0'; . '.\Laptop_Inspector_Fixed.ps1'}"
)

echo.
echo Script execution completed.
echo Check your Desktop for the HTML report.
echo.
pause
