@echo off
setlocal EnableDelayedExpansion

:: Check if running as Administrator
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: This script must be run as Administrator.
    echo Reason: Administrative privileges are required to install software via Chocolatey.
    echo Solution: Right-click this file and select "Run as Administrator".
    pause
    exit /b 1
)

:: Verify PowerShell is installed and get version
where powershell >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: PowerShell is not installed or not found in PATH.
    echo Reason: The 'powershell' command could not be located in your system PATH.
    echo Solution: Ensure PowerShell is installed and added to the system PATH environment variable.
    echo You can download PowerShell from https://docs.microsoft.com/en-us/powershell/.
    pause
    exit /b 1
)

:: Check PowerShell version
for /f "tokens=*" %%i in ('powershell -NoProfile -Command "$PSVersionTable.PSVersion.Major"') do set PSVersion=%%i
if %errorlevel% neq 0 (
    echo ERROR: Failed to determine PowerShell version.
    echo Reason: PowerShell may be corrupted or not functioning correctly.
    echo Solution: Reinstall PowerShell or verify its configuration.
    pause
    exit /b 1
)
if %PSVersion% LSS 5 (
    echo WARNING: PowerShell version %PSVersion% detected. Version 5 or higher is recommended.
    echo Reason: Older versions may have compatibility issues with Chocolatey scripts.
    echo Solution: Consider upgrading PowerShell for better compatibility.
)

:: Check if the PowerShell script exists
if not exist "%~dp0PerdangaSoftwareSolutions.ps1" (
    echo ERROR: PerdangaSoftwareSolutions.ps1 not found in the same directory.
    echo Reason: The script file is missing from %~dp0.
    echo Solution: Ensure PerdangaSoftwareSolutions.ps1 is placed in the same directory as this batch file.
    pause
    exit /b 1
)

:: Launch PowerShell script in a new PowerShell console
echo Launching PowerShell script...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process powershell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%~dp0PerdangaSoftwareSolutions.ps1""' -Verb RunAs" 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Failed to launch PowerShell script.
    echo Reason: Possible issues include corrupted script, insufficient permissions, or execution policy restrictions.
    echo Solution: Verify that PerdangaSoftwareSolutions.ps1 is not corrupted, check execution policy (try 'Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass'), or ensure you have sufficient permissions.
    echo Error Code: %errorlevel%
    pause
    exit /b %errorlevel%
)

echo PowerShell script launched successfully.
exit /b 0
