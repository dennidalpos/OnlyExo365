#!/usr/bin/env pwsh
#Requires -Version 7.0

<#
.SYNOPSIS
    Quick build and test script for ExchangeAdmin

.DESCRIPTION
    Cleans old builds, compiles, publishes, and launches the application
#>

param(
    [switch]$SkipClean,
    [switch]$SkipBuild,
    [switch]$LaunchOnly
)

$ErrorActionPreference = "Stop"

# Colors
$Red = "`e[31m"
$Green = "`e[32m"
$Yellow = "`e[33m"
$Blue = "`e[34m"
$Reset = "`e[0m"

function Write-Step {
    param([string]$Message)
    Write-Host "${Blue}==>${Reset} ${Message}"
}

function Write-Success {
    param([string]$Message)
    Write-Host "${Green}✓${Reset} ${Message}"
}

function Write-Error {
    param([string]$Message)
    Write-Host "${Red}✗${Reset} ${Message}"
}

function Write-Warning {
    param([string]$Message)
    Write-Host "${Yellow}⚠${Reset} ${Message}"
}

# Get script directory and project root
$ScriptDir = Split-Path -Parent $PSCommandPath
$ProjectRoot = Split-Path -Parent $ScriptDir
$PublishDir = Join-Path $ProjectRoot "artifacts\publish"
$PresentationExe = Join-Path $PublishDir "ExchangeAdmin.Presentation.exe"
$WorkerExe = Join-Path $PublishDir "ExchangeAdmin.Worker.exe"

Write-Host ""
Write-Host "${Blue}╔════════════════════════════════════════╗${Reset}"
Write-Host "${Blue}║${Reset}  ExchangeAdmin Quick Build & Test  ${Blue}║${Reset}"
Write-Host "${Blue}╔════════════════════════════════════════╗${Reset}"
Write-Host ""

# Change to project root
Set-Location $ProjectRoot

# Step 1: Clean old builds
if (-not $SkipClean -and -not $LaunchOnly) {
    Write-Step "Cleaning old builds..."

    $cleanPaths = @(
        "artifacts\publish",
        "src\ExchangeAdmin.Presentation\bin\Release",
        "src\ExchangeAdmin.Presentation\obj\Release",
        "src\ExchangeAdmin.Worker\bin\Release",
        "src\ExchangeAdmin.Worker\obj\Release",
        "src\ExchangeAdmin.Infrastructure\bin\Release",
        "src\ExchangeAdmin.Infrastructure\obj\Release",
        "src\ExchangeAdmin.Application\bin\Release",
        "src\ExchangeAdmin.Application\obj\Release",
        "src\ExchangeAdmin.Domain\bin\Release",
        "src\ExchangeAdmin.Domain\obj\Release",
        "src\ExchangeAdmin.Contracts\bin\Release",
        "src\ExchangeAdmin.Contracts\obj\Release"
    )

    $cleanedCount = 0
    foreach ($path in $cleanPaths) {
        $fullPath = Join-Path $ProjectRoot $path
        if (Test-Path $fullPath) {
            Remove-Item $fullPath -Recurse -Force -ErrorAction SilentlyContinue
            $cleanedCount++
        }
    }

    Write-Success "Cleaned $cleanedCount directories"
}

# Step 2: Build
if (-not $SkipBuild -and -not $LaunchOnly) {
    Write-Step "Building solution..."

    $buildResult = dotnet build -c Release --nologo --verbosity quiet 2>&1

    if ($LASTEXITCODE -ne 0) {
        Write-Error "Build failed!"
        Write-Host $buildResult
        exit 1
    }

    Write-Success "Build completed"
}

# Step 3: Publish
if (-not $LaunchOnly) {
    Write-Step "Publishing applications..."

    # Publish Presentation
    $publishResult = dotnet publish -c Release src/ExchangeAdmin.Presentation/ExchangeAdmin.Presentation.csproj `
        -o $PublishDir --nologo --verbosity quiet --no-build 2>&1

    if ($LASTEXITCODE -ne 0) {
        Write-Error "Publish failed for Presentation!"
        Write-Host $publishResult
        exit 1
    }

    # Publish Worker
    $publishResult = dotnet publish -c Release src/ExchangeAdmin.Worker/ExchangeAdmin.Worker.csproj `
        -o $PublishDir --nologo --verbosity quiet --no-build 2>&1

    if ($LASTEXITCODE -ne 0) {
        Write-Error "Publish failed for Worker!"
        Write-Host $publishResult
        exit 1
    }

    Write-Success "Published to: $PublishDir"
}

# Step 4: Verify executables exist
Write-Step "Verifying executables..."

if (-not (Test-Path $PresentationExe)) {
    Write-Error "Presentation executable not found at: $PresentationExe"
    exit 1
}

if (-not (Test-Path $WorkerExe)) {
    Write-Error "Worker executable not found at: $WorkerExe"
    exit 1
}

Write-Success "Executables verified"

# Step 5: Launch
Write-Host ""
Write-Host "${Green}╔════════════════════════════════════════╗${Reset}"
Write-Host "${Green}║${Reset}       Launching Application        ${Green}║${Reset}"
Write-Host "${Green}╚════════════════════════════════════════╝${Reset}"
Write-Host ""
Write-Host "Presentation: ${Yellow}$PresentationExe${Reset}"
Write-Host "Worker:       ${Yellow}$WorkerExe${Reset}"
Write-Host ""
Write-Warning "Press Ctrl+C in this terminal to stop tracking the application"
Write-Host ""

# Launch the application
try {
    $process = Start-Process -FilePath $PresentationExe -WorkingDirectory $PublishDir -PassThru

    Write-Success "Application launched (PID: $($process.Id))"
    Write-Host ""
    Write-Host "${Blue}Waiting for application to exit...${Reset}"
    Write-Host "${Yellow}(The Worker console window will open separately)${Reset}"

    # Wait for the process to exit
    $process.WaitForExit()

    Write-Host ""
    Write-Success "Application closed (Exit Code: $($process.ExitCode))"
}
catch {
    Write-Error "Failed to launch application: $_"
    exit 1
}
