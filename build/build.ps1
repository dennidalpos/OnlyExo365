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
    [string[]]$RuntimeIdentifiers = @("win-x64", "win-x86"),

    [switch]$LockedMode = $false,

    [string]$ExportDirPath,

    [string]$ImportDirPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepositoryRoot = Split-Path -Parent $ScriptDir

. (Join-Path $RepositoryRoot "scripts\helpers\common.ps1")

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

    $bootstrapPolicyPath = Join-Path $RepositoryRoot "src\ExchangeAdmin.Worker\Data\PowerShellModuleBootstrapPolicy.json"

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

function Normalize-Path {
    param([string]$PathValue)

    if ([string]::IsNullOrWhiteSpace($PathValue)) {
        return $null
    }

    try {
        return [System.IO.Path]::GetFullPath($PathValue).TrimEnd('\')
    }
    catch {
        return $PathValue.TrimEnd('\')
    }
}

function Test-PathWithinRoot {
    param(
        [string]$CandidatePath,
        [string]$RootPath
    )

    $normalizedCandidate = Normalize-Path -PathValue $CandidatePath
    $normalizedRoot = Normalize-Path -PathValue $RootPath

    if ([string]::IsNullOrWhiteSpace($normalizedCandidate) -or [string]::IsNullOrWhiteSpace($normalizedRoot)) {
        return $false
    }

    return $normalizedCandidate.StartsWith($normalizedRoot, [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-ArtifactExchangeAdminProcesses {
    param([string]$RootPath)

    $targetNames = @("ExchangeAdmin.Presentation.exe", "ExchangeAdmin.Worker.exe")
    $processes = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object {
        $targetNames -contains $_.Name -and
        -not [string]::IsNullOrWhiteSpace($_.ExecutablePath) -and
        (Test-PathWithinRoot -CandidatePath $_.ExecutablePath -RootPath $RootPath)
    }

    return @($processes | Sort-Object ProcessId -Unique)
}

function Stop-ArtifactExchangeAdminProcesses {
    param([string]$RootPath)

    $processes = @(Get-ArtifactExchangeAdminProcesses -RootPath $RootPath)
    if ($processes.Count -eq 0) {
        return $false
    }

    Write-Warn "Detected running ExchangeAdmin process(es) from artifacts. Stopping them before clean."

    foreach ($process in $processes) {
        Write-Info "Stopping PID $($process.ProcessId): $($process.ExecutablePath)"
        try {
            Stop-Process -Id $process.ProcessId -Force -ErrorAction Stop
        }
        catch {
            Write-Warn "Unable to stop PID $($process.ProcessId): $($_.Exception.Message)"
        }
    }

    Start-Sleep -Milliseconds 750
    return $true
}

function Remove-DirectoryRobust {
    param(
        [string]$Path,
        [string]$Description,
        [string]$ArtifactRootForProcessStop = $null,
        [int]$RetryCount = 3
    )

    if (-not (Test-Path $Path)) {
        return
    }

    $stoppedProcesses = $false

    for ($attempt = 1; $attempt -le $RetryCount; $attempt++) {
        try {
            Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
            Write-Success "Removed ${Description}: $Path"
            return
        }
        catch {
            if (-not $stoppedProcesses -and -not [string]::IsNullOrWhiteSpace($ArtifactRootForProcessStop)) {
                $stoppedProcesses = Stop-ArtifactExchangeAdminProcesses -RootPath $ArtifactRootForProcessStop
            }

            if ($attempt -lt $RetryCount) {
                Write-Warn "Retrying cleanup for $Description after failure: $($_.Exception.Message)"
                Start-Sleep -Milliseconds (500 * $attempt)
                continue
            }

            $hint = if ($stoppedProcesses) {
                "A lock is still active after stopping ExchangeAdmin artifact processes."
            }
            else {
                "No ExchangeAdmin process from artifacts was found to stop automatically."
            }

            Stop-WithError "Unable to remove $Description at '$Path'. $hint Close external file handles and retry. Root cause: $($_.Exception.Message)"
        }
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
$ResolvedRuntimeIdentifiers = Resolve-RuntimeIdentifiers -RequestedRuntimeIdentifiers $RuntimeIdentifiers -DefaultRuntimeIdentifiers @("win-x64", "win-x86")

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " ExchangeAdmin Build Script" -ForegroundColor Cyan
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

    if (Test-Path $OutputDir) {
        Remove-DirectoryRobust -Path $OutputDir -Description "artifacts output" -ArtifactRootForProcessStop $OutputDir
    }
    else {
        Write-Info "Output directory doesn't exist, skipping"
    }

    Invoke-DotNet -Command "clean" -Arguments @($SolutionFile, "-c", $Configuration, "--verbosity", "minimal") -ErrorMessage "Clean failed"
    Write-Success "Solution cleaned"
}

Initialize-ArtifactsLayout -RepositoryRoot $RepositoryRoot | Out-Null

Write-Step "Restoring NuGet packages"

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
        Remove-DirectoryRobust -Path $PublishDir -Description "publish output"
    }

    New-Item -Path $PublishDir -ItemType Directory -Force | Out-Null
    Write-Info "Prepared publish directory root: $PublishDir"

    foreach ($pathInfo in @(@{ Path = $ExportDir; Label = "exports" }, @{ Path = $ImportDir; Label = "imports" })) {
        if (-not (Test-Path $pathInfo.Path)) {
            New-Item -Path $pathInfo.Path -ItemType Directory -Force | Out-Null
            Write-Info "Created $($pathInfo.Label) directory: $($pathInfo.Path)"
        }
    }

    $presentationProject = Join-Path $RepositoryRoot "src\ExchangeAdmin.Presentation\ExchangeAdmin.Presentation.csproj"
    $workerProject = Join-Path $RepositoryRoot "src\ExchangeAdmin.Worker\ExchangeAdmin.Worker.csproj"

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

        Write-Info "Publishing ExchangeAdmin.Presentation for $runtimeIdentifier..."
        Invoke-DotNet -Command "publish" -Arguments (@($presentationProject) + $publishArgs) -ErrorMessage "Publish of Presentation failed"
        Write-Success "ExchangeAdmin.Presentation published for $runtimeIdentifier"

        Write-Info "Publishing ExchangeAdmin.Worker for $runtimeIdentifier..."
        Invoke-DotNet -Command "publish" -Arguments (@($workerProject) + $publishArgs) -ErrorMessage "Publish of Worker failed"
        Write-Success "ExchangeAdmin.Worker published for $runtimeIdentifier"
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
    Write-Host "  cd `"$((Get-RuntimePublishPath -RepositoryRoot $RepositoryRoot -RuntimeIdentifier ($ResolvedRuntimeIdentifiers[0])))`""
    Write-Host "  .\ExchangeAdmin.Presentation.exe"
    Write-Host ""
    Write-Host "Runtime prerequisites:" -ForegroundColor Yellow
    Write-Host "  1. PowerShell 7+ (pwsh.exe in PATH)"
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
