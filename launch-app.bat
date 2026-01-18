@echo off
REM Launch ExchangeAdmin without rebuilding
REM Use this for quick testing after you've already built once

echo.
echo ═══════════════════════════════════════
echo   Launching ExchangeAdmin (No Build)
echo ═══════════════════════════════════════
echo.

set PUBLISH_DIR=%~dp0artifacts\publish
set EXE_PATH=%PUBLISH_DIR%\ExchangeAdmin.Presentation.exe

if not exist "%EXE_PATH%" (
    echo ERROR: Application not found at:
    echo %EXE_PATH%
    echo.
    echo Run quick-test.bat first to build the application.
    pause
    exit /b 1
)

echo Launching: %EXE_PATH%
echo.

start "" "%EXE_PATH%"

echo Application launched.
echo The Worker console window will open separately.
echo.
pause
