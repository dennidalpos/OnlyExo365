@echo off
REM Quick build and test launcher for ExchangeAdmin
REM Requires PowerShell 7+ (pwsh.exe)

where pwsh.exe >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Error: PowerShell 7+ is required but not found in PATH
    echo Install from: https://github.com/PowerShell/PowerShell/releases
    pause
    exit /b 1
)

pwsh.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0quick-test.ps1" %*
