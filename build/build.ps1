#Requires -Version 7.0

<#
.SYNOPSIS
    Build script for ExchangeAdmin solution.

.DESCRIPTION
    Builds and publishes the ExchangeAdmin WPF application and out-of-process PowerShell worker.
    Implements fail-fast semantics: any error stops the build immediately.

.PARAMETER Configuration
    Build configuration: Debug or Release. Default is Release.

.PARAMETER Clean
    Clean before building. Removes artifacts directory and runs dotnet clean.

.PARAMETER Publish
    Publish the application after building.

.PARAMETER SelfContained
    Create self-contained deployment (includes .NET runtime).
    Output will be larger (~150MB) but doesn't require .NET 8 on target machine.

.PARAMETER Verbose
    Show detailed output from dotnet commands.

.EXAMPLE
    .\build.ps1 -Configuration Release -Publish
    # Standard release build with publish

.EXAMPLE
    .\build.ps1 -Clean -Publish -SelfContained
    # Clean build with self-contained publish

.EXAMPLE
    .\build.ps1 -Configuration Debug -Verbose
    # Debug build with verbose output

.NOTES
    Prerequisites:
    - .NET 8 SDK
    - PowerShell 7+
    - Windows (WPF is Windows-only)

    For runtime:
    - PowerShell 7+ (for worker process)
    - ExchangeOnlineManagement module (Install-Module ExchangeOnlineManagement)
#>

[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",

    [switch]$Clean = $true,

    [switch]$Publish = $true,

    [switch]$SelfContained = $true
)

# Strict error handling - fail fast
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

# Paths
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$SolutionDir = Split-Path -Parent $ScriptDir
$SolutionFile = Join-Path $SolutionDir "ExchangeAdmin.sln"
$OutputDir = Join-Path $SolutionDir "artifacts"
$PublishDir = Join-Path $OutputDir "publish"

# Timestamp for logging
$BuildStartTime = Get-Date

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host ">> $Message" -ForegroundColor Cyan
}

function Write-Success {
    param([string]$Message)
    Write-Host "   [OK] $Message" -ForegroundColor Green
}

function Write-Info {
    param([string]$Message)
    Write-Host "   $Message" -ForegroundColor Gray
}

function Write-Warn {
    param([string]$Message)
    Write-Host "   [WARN] $Message" -ForegroundColor Yellow
}

function Stop-WithError {
    param([string]$Message, [int]$ExitCode = 1)
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "[FAILED] $Message" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    exit $ExitCode
}

function Invoke-DotNet {
    param(
        [string]$Command,
        [string[]]$Arguments,
        [string]$ErrorMessage
    )

    $allArgs = @($Command) + $Arguments

    if ($VerbosePreference -eq 'Continue') {
        Write-Info "dotnet $($allArgs -join ' ')"
        & dotnet @allArgs
    }
    else {
        & dotnet @allArgs 2>&1 | ForEach-Object {
            if ($_ -match 'error') {
                Write-Host "   $_" -ForegroundColor Red
            }
            elseif ($_ -match 'warning') {
                Write-Host "   $_" -ForegroundColor Yellow
            }
            elseif ($VerbosePreference -eq 'Continue') {
                Write-Host "   $_" -ForegroundColor Gray
            }
        }
    }

    if ($LASTEXITCODE -ne 0) {
        Stop-WithError "$ErrorMessage (exit code: $LASTEXITCODE)" $LASTEXITCODE
    }
}

# Banner
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " ExchangeAdmin Build Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Configuration : $Configuration"
Write-Host "Solution      : $SolutionFile"
Write-Host "Output        : $OutputDir"
Write-Host "Publish       : $($Publish.IsPresent)"
Write-Host "Self-contained: $($SelfContained.IsPresent)"
Write-Host "Started at    : $($BuildStartTime.ToString('HH:mm:ss'))"

# Verify solution exists
if (-not (Test-Path $SolutionFile)) {
    Stop-WithError "Solution file not found: $SolutionFile"
}

# Verify .NET SDK
Write-Step "Checking prerequisites"

try {
    $dotnetVersion = & dotnet --version 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "dotnet not found"
    }
    Write-Success ".NET SDK version: $dotnetVersion"

    # Verify it's .NET 8+
    $majorVersion = [int]($dotnetVersion.Split('.')[0])
    if ($majorVersion -lt 8) {
        Stop-WithError ".NET 8 SDK or later is required. Found: $dotnetVersion"
    }
}
catch {
    Stop-WithError ".NET SDK not found. Please install .NET 8 SDK from https://dotnet.microsoft.com/download"
}

# Verify we're on Windows (WPF requirement)
if ($PSVersionTable.Platform -and $PSVersionTable.Platform -ne 'Win32NT') {
    Stop-WithError "This project requires Windows (WPF is Windows-only)"
}

# Clean
if ($Clean) {
    Write-Step "Cleaning"

    if (Test-Path $OutputDir) {
        Remove-Item -Path $OutputDir -Recurse -Force -ErrorAction Stop
        Write-Success "Removed: $OutputDir"
    }
    else {
        Write-Info "Output directory doesn't exist, skipping"
    }

    Invoke-DotNet -Command "clean" -Arguments @($SolutionFile, "-c", $Configuration, "--verbosity", "minimal") `
                  -ErrorMessage "Clean failed"
    Write-Success "Solution cleaned"
}

# Restore
Write-Step "Restoring NuGet packages"

$restoreArgs = @($SolutionFile, "--verbosity", "minimal")

# Per self-contained, serve restore con runtime identifier
if ($SelfContained) {
    $restoreArgs += "-r", "win-x64"
    Write-Info "Restoring for runtime: win-x64"
}

Invoke-DotNet -Command "restore" -Arguments $restoreArgs `
              -ErrorMessage "Package restore failed"
Write-Success "Packages restored"

# Build
Write-Step "Building solution"

$buildArgs = @(
    $SolutionFile
    "-c", $Configuration
    "--no-restore"
    "-warnaserror:nullable"
    "-p:TreatWarningsAsErrors=false"
)

if ($VerbosePreference -ne 'Continue') {
    $buildArgs += "--verbosity", "minimal"
}

Invoke-DotNet -Command "build" -Arguments $buildArgs -ErrorMessage "Build failed"
Write-Success "Build succeeded"

# Publish
if ($Publish) {
    Write-Step "Publishing applications"

    # Create publish directory
    if (-not (Test-Path $PublishDir)) {
        New-Item -Path $PublishDir -ItemType Directory -Force | Out-Null
        Write-Info "Created: $PublishDir"
    }

    $publishArgs = @(
        "-c", $Configuration
        "-o", $PublishDir
    )

    if ($SelfContained) {
        # Self-contained richiede build con RID, non possiamo usare --no-build
        $publishArgs += "--self-contained", "true"
        $publishArgs += "-r", "win-x64"
        $publishArgs += "-p:PublishSingleFile=false"
        Write-Info "Mode: Self-contained (win-x64)"
    }
    else {
        # Framework-dependent può riusare il build precedente
        $publishArgs += "--no-build"
        $publishArgs += "--self-contained", "false"
        Write-Info "Mode: Framework-dependent"
    }

    if ($VerbosePreference -ne 'Continue') {
        $publishArgs += "--verbosity", "minimal"
    }

    # Publish main application
    Write-Info "Publishing ExchangeAdmin.Presentation..."
    $presentationProject = Join-Path $SolutionDir "src\ExchangeAdmin.Presentation\ExchangeAdmin.Presentation.csproj"

    if (-not (Test-Path $presentationProject)) {
        Stop-WithError "Presentation project not found: $presentationProject"
    }

    Invoke-DotNet -Command "publish" -Arguments (@($presentationProject) + $publishArgs) `
                  -ErrorMessage "Publish of Presentation failed"
    Write-Success "ExchangeAdmin.Presentation published"

    # Publish worker
    Write-Info "Publishing ExchangeAdmin.Worker..."
    $workerProject = Join-Path $SolutionDir "src\ExchangeAdmin.Worker\ExchangeAdmin.Worker.csproj"

    if (-not (Test-Path $workerProject)) {
        Stop-WithError "Worker project not found: $workerProject"
    }

    Invoke-DotNet -Command "publish" -Arguments (@($workerProject) + $publishArgs) `
                  -ErrorMessage "Publish of Worker failed"
    Write-Success "ExchangeAdmin.Worker published"

    # List published files
    $publishedFiles = Get-ChildItem -Path $PublishDir -Filter "*.exe" | Select-Object -ExpandProperty Name
    Write-Info "Published executables: $($publishedFiles -join ', ')"

    # Calculate size
    $publishSize = (Get-ChildItem -Path $PublishDir -Recurse | Measure-Object -Property Length -Sum).Sum / 1MB
    Write-Info "Total size: $([math]::Round($publishSize, 2)) MB"
}

# Summary
$BuildEndTime = Get-Date
$BuildDuration = $BuildEndTime - $BuildStartTime

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " BUILD COMPLETED SUCCESSFULLY" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Duration: $($BuildDuration.TotalSeconds.ToString('F1')) seconds" -ForegroundColor Gray

if ($Publish) {
    Write-Host ""
    Write-Host "Published to: $PublishDir" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "To run the application:" -ForegroundColor Yellow
    Write-Host "  cd `"$PublishDir`""
    Write-Host "  .\ExchangeAdmin.Presentation.exe"
    Write-Host ""
    Write-Host "Runtime prerequisites:" -ForegroundColor Yellow
    Write-Host "  1. PowerShell 7+ (pwsh.exe in PATH)"
    Write-Host "  2. ExchangeOnlineManagement module:"
    Write-Host "     Install-Module ExchangeOnlineManagement -Scope CurrentUser"

    if (-not $SelfContained) {
        Write-Host "  3. .NET 8 Runtime (framework-dependent build)"
    }
}

Write-Host ""
exit 0
