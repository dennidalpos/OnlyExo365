#Requires -Version 7.0

[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",

    [switch]$Clean = $true,

    [switch]$Publish = $true,

    [switch]$SelfContained = $false,

    [Alias("RuntimeIdentifier")]
    [ValidateNotNullOrEmpty()]
    [string[]]$RuntimeIdentifiers = @("win-x64"),

    [switch]$LockedMode = $false,

    [switch]$NoRestore = $false,

    [string]$ExportDirPath,

    [string]$ImportDirPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepositoryRoot = Split-Path -Parent $ScriptDir

. (Join-Path $RepositoryRoot "scripts\internal\common.ps1")

function Test-IsCiBuild {
    $ciMarkers = @(
        $env:CI,
        $env:TF_BUILD,
        $env:GITHUB_ACTIONS,
        $env:BUILD_BUILDID,
        $env:JENKINS_URL
    )

    foreach ($marker in $ciMarkers) {
        if (-not [string]::IsNullOrWhiteSpace($marker) -and $marker -ne "false" -and $marker -ne "0") {
            return $true
        }
    }

    return $false
}

function Get-BootstrapModuleVersion {
    param([string]$ModuleName)

    $bootstrapPolicyPath = Join-Path $RepositoryRoot "src\OnlyExo365.Contracts\Data\PowerShellModuleBootstrapPolicy.json"

    if (-not (Test-Path $bootstrapPolicyPath)) {
        return $null
    }

    try {
        $policy = Get-Content -Raw $bootstrapPolicyPath | ConvertFrom-Json -ErrorAction Stop
        return ($policy.modules | Where-Object { $_.moduleName -eq $ModuleName } | Select-Object -First 1).requiredVersion
    }
    catch {
        Write-Warn "Unable to read PowerShell module bootstrap policy: $($_.Exception.Message)"
        return $null
    }
}

function Invoke-DotNet {
    param(
        [string]$Command,
        [string[]]$Arguments,
        [string]$ErrorMessage
    )

    $allArgs = @($Command) + $Arguments

    if ($VerbosePreference -eq "Continue") {
        Write-Info "dotnet $($allArgs -join ' ')"
        & dotnet @allArgs
    }
    else {
        & dotnet @allArgs 2>&1 | ForEach-Object {
            if ($_ -match "error") {
                Write-Host "   $_" -ForegroundColor Red
            }
            elseif ($_ -match "warning") {
                Write-Host "   $_" -ForegroundColor Yellow
            }
        }
    }

    if ($LASTEXITCODE -ne 0) {
        Stop-WithError "$ErrorMessage (exit code: $LASTEXITCODE)" $LASTEXITCODE
    }
}

$IsCiBuild = Test-IsCiBuild
$SelfContainedMode = [bool]$SelfContained
$SolutionFile = Get-SolutionPath -RepositoryRoot $RepositoryRoot
$OutputDir = Join-Path $RepositoryRoot "artifacts"
$BuildArtifactsDir = Get-BuildArtifactsPath -RepositoryRoot $RepositoryRoot
$PublishDir = Get-PublishArtifactsPath -RepositoryRoot $RepositoryRoot
$ExportDir = if ([string]::IsNullOrWhiteSpace($ExportDirPath)) { Join-Path $OutputDir "exports" } else { Resolve-RepositoryPath -RepositoryRoot $RepositoryRoot -PathValue $ExportDirPath }
$ImportDir = if ([string]::IsNullOrWhiteSpace($ImportDirPath)) { Join-Path $OutputDir "imports" } else { Resolve-RepositoryPath -RepositoryRoot $RepositoryRoot -PathValue $ImportDirPath }
$BuildStartTime = Get-Date
$ResolvedRuntimeIdentifiers = Resolve-RuntimeIdentifiers -RequestedRuntimeIdentifiers $RuntimeIdentifiers -DefaultRuntimeIdentifiers @("win-x64")
$PrimaryRuntimeIdentifier = @($ResolvedRuntimeIdentifiers)[0]

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " OnlyExo365 Build Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Configuration : $Configuration"
Write-Host "Solution      : $SolutionFile"
Write-Host "Output        : $OutputDir"
Write-Host "Build artifacts: $BuildArtifactsDir"
Write-Host "Exports       : $ExportDir"
Write-Host "Imports       : $ImportDir"
Write-Host "Publish       : $($Publish.IsPresent)"
Write-Host "Self-contained: $SelfContainedMode"
Write-Host "Runtimes      : $($ResolvedRuntimeIdentifiers -join ', ')"
Write-Host "Locked mode   : $($LockedMode.IsPresent)"
Write-Host "Restore       : $(-not $NoRestore.IsPresent)"
Write-Host "CI            : $IsCiBuild"
Write-Host "Started at    : $($BuildStartTime.ToString('HH:mm:ss'))"

if (-not (Test-Path $SolutionFile)) {
    Stop-WithError "Solution file not found: $SolutionFile"
}

Write-Step "Checking prerequisites"

try {
    Assert-WindowsPlatform
    $sdkSpecification = Assert-DotNetSdkPinnedVersion -RepositoryRoot $RepositoryRoot
    Write-Success ".NET SDK version: $($sdkSpecification.Version)"
}
catch {
    Stop-WithError $_.Exception.Message
}

if ($Clean) {
    Write-Step "Cleaning"

    Invoke-RepositoryPowerShellScript -ScriptPath (Join-Path $RepositoryRoot 'build/clean.ps1') `
        -Arguments @('-SkipDotNetClean') -ErrorMessage 'Generated output cleanup failed'

    Invoke-DotNet -Command "clean" -Arguments @($SolutionFile, "-c", $Configuration, "--verbosity", "minimal") -ErrorMessage "Clean failed"
    Write-Success "Solution cleaned"
}

Initialize-ArtifactsLayout -RepositoryRoot $RepositoryRoot | Out-Null

Write-Step "Restoring NuGet packages"

if (-not $NoRestore) {
    $restoreArgs = @(
        $SolutionFile,
        "--verbosity", "minimal",
        "--artifacts-path", $BuildArtifactsDir
    )

    if ($LockedMode) {
        $restoreArgs += "--locked-mode"
        Write-Info "NuGet restore locked mode enabled"
    }

    Write-Info "Restoring with project runtime graph: $($ResolvedRuntimeIdentifiers -join ', ')"
    Invoke-DotNet -Command "restore" -Arguments $restoreArgs -ErrorMessage "Package restore failed"

    Write-Success "Packages restored"
}
else {
    Write-Info "Restore skipped by -NoRestore"
}

Write-Step "Building solution"

$buildArgs = @(
    $SolutionFile
    "-c", $Configuration
    "--no-restore"
    "--artifacts-path", $BuildArtifactsDir
    "-warnaserror:nullable"
)

if ($VerbosePreference -ne "Continue") {
    $buildArgs += "--verbosity", "minimal"
}

Invoke-DotNet -Command "build" -Arguments $buildArgs -ErrorMessage "Build failed"
Write-Success "Build succeeded"

if ($Publish) {
    Write-Step "Publishing applications"

    if (Test-Path $PublishDir) {
        $null = Assert-RepositoryCleanupPath -RepositoryRoot $RepositoryRoot -PathValue $PublishDir `
            -TrackedPaths @(Get-TrackedRepositoryPaths -RepositoryRoot $RepositoryRoot)
        Remove-Item -LiteralPath $PublishDir -Recurse -Force -ErrorAction Stop
    }

    New-Item -Path $PublishDir -ItemType Directory -Force | Out-Null
    Write-Info "Prepared publish directory root: $PublishDir"

    foreach ($pathInfo in @(@{ Path = $ExportDir; Label = "exports" }, @{ Path = $ImportDir; Label = "imports" })) {
        if (-not (Test-Path $pathInfo.Path)) {
            New-Item -Path $pathInfo.Path -ItemType Directory -Force | Out-Null
            Write-Info "Created $($pathInfo.Label) directory: $($pathInfo.Path)"
        }
    }

    $presentationProject = Join-Path $RepositoryRoot "src\OnlyExo365.Shell\OnlyExo365.Shell.csproj"
    $workerProject = Join-Path $RepositoryRoot "src\OnlyExo365.Worker\OnlyExo365.Worker.csproj"

    if (-not (Test-Path $presentationProject)) {
        Stop-WithError "Presentation project not found: $presentationProject"
    }

    if (-not (Test-Path $workerProject)) {
        Stop-WithError "Worker project not found: $workerProject"
    }

    foreach ($runtimeIdentifier in $ResolvedRuntimeIdentifiers) {
        $runtimePublishDir = Get-RuntimePublishPath -RepositoryRoot $RepositoryRoot -RuntimeIdentifier $runtimeIdentifier
        New-Item -Path $runtimePublishDir -ItemType Directory -Force | Out-Null

        $publishArgs = @(
            "-c", $Configuration
            "-o", $runtimePublishDir
            "--artifacts-path", $BuildArtifactsDir
            "--no-restore"
            "-r", $runtimeIdentifier
        )

        if ($SelfContainedMode) {
            $publishArgs += "--self-contained", "true"
            $publishArgs += "-p:PublishSingleFile=false"
            Write-Info "Mode: Self-contained ($runtimeIdentifier)"
        }
        else {
            $publishArgs += "--self-contained", "false"
            Write-Info "Mode: Framework-dependent ($runtimeIdentifier)"
        }

        if ($VerbosePreference -ne "Continue") {
            $publishArgs += "--verbosity", "minimal"
        }

        Write-Info "Publishing OnlyExo365.Shell for $runtimeIdentifier..."
        Invoke-DotNet -Command "publish" -Arguments (@($presentationProject) + $publishArgs) -ErrorMessage "Publish of Presentation failed"
        Write-Success "OnlyExo365.Shell published for $runtimeIdentifier"

        Write-Info "Publishing OnlyExo365.Worker for $runtimeIdentifier..."
        Invoke-DotNet -Command "publish" -Arguments (@($workerProject) + $publishArgs) -ErrorMessage "Publish of Worker failed"
        Write-Success "OnlyExo365.Worker published for $runtimeIdentifier"
    }

    $publishedFiles = Get-ChildItem -Path $PublishDir -Recurse -Filter "*.exe" | Select-Object -ExpandProperty FullName
    $publishSize = (Get-ChildItem -Path $PublishDir -Recurse | Measure-Object -Property Length -Sum).Sum / 1MB
    Write-Info "Published executables: $($publishedFiles -join ', ')"
    Write-Info "Total size: $([math]::Round($publishSize, 2)) MB"
}

$BuildEndTime = Get-Date
$BuildDuration = $BuildEndTime - $BuildStartTime

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " BUILD COMPLETED SUCCESSFULLY" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Duration: $($BuildDuration.TotalSeconds.ToString('F1')) seconds" -ForegroundColor Gray

if ($Publish) {
    $exchangeModuleRequiredVersion = Get-BootstrapModuleVersion -ModuleName "ExchangeOnlineManagement"
    Write-Host ""
    Write-Host "Published to: $PublishDir" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "To run the application:" -ForegroundColor Yellow
    Write-Host "  cd `"$((Get-RuntimePublishPath -RepositoryRoot $RepositoryRoot -RuntimeIdentifier $PrimaryRuntimeIdentifier))`""
    Write-Host "  .\OnlyExo365.Shell.exe"
    Write-Host ""
    Write-Host "Runtime prerequisites:" -ForegroundColor Yellow
    Write-Host "  1. PowerShell 7.6+ (pwsh.exe in PATH) for interactive Exchange bootstrap"
    Write-Host "  2. ExchangeOnlineManagement module:"
    if ([string]::IsNullOrWhiteSpace($exchangeModuleRequiredVersion)) {
        Write-Host "     Install-Module ExchangeOnlineManagement -Repository PSGallery -Scope CurrentUser"
    }
    else {
        Write-Host "     Install-Module ExchangeOnlineManagement -Repository PSGallery -RequiredVersion $exchangeModuleRequiredVersion -Scope CurrentUser -AllowClobber"
    }

    if (-not $SelfContainedMode) {
        Write-Host "  3. .NET 10 Runtime (framework-dependent build)"
    }
}

