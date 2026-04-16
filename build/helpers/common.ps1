#Requires -Version 7.0

function Write-Step {
    param([string]$Message)

    Write-Host ""
    Write-Host ">> $Message" -ForegroundColor Cyan
}

function Write-Info {
    param([string]$Message)

    Write-Host "   $Message" -ForegroundColor Gray
}

function Write-Success {
    param([string]$Message)

    Write-Host "   [OK] $Message" -ForegroundColor Green
}

function Write-Warn {
    param([string]$Message)

    Write-Host "   [WARN] $Message" -ForegroundColor Yellow
}

function Stop-WithError {
    param(
        [string]$Message,
        [int]$ExitCode = 1
    )

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "[FAILED] $Message" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    exit $ExitCode
}
