@echo off
echo Starting Skype CSV to VCF Converter...
echo.

REM Check if PowerShell is available
powershell -Command "Get-Host" >nul 2>&1
if %errorlevel% neq 0 (
    echo PowerShell is not available on this system.
    echo Please install PowerShell or run the .ps1 file directly.
    pause
    exit /b 1
)

REM Run the PowerShell script
powershell -ExecutionPolicy Bypass -File "%~dp0Convert-SkypeCSVToVCF.ps1"

REM Keep window open if there was an error
if %errorlevel% neq 0 (
    echo.
    echo Script completed with errors.
    pause
)
