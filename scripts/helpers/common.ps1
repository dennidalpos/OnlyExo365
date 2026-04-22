#Requires -Version 7.0

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-RepositoryRoot {
    param([string]$ScriptRoot)

    return (Split-Path -Parent $ScriptRoot)
}

function Get-SolutionPath {
    param([string]$RepositoryRoot)

    return (Join-Path $RepositoryRoot "OnlyExo365.sln")
}

function Get-BuildArtifactsPath {
    param([string]$RepositoryRoot)

    return (Join-Path $RepositoryRoot "artifacts\\build")
}

function Get-ApplicationVersion {
    param([string]$RepositoryRoot)

    $propsPath = Join-Path $RepositoryRoot "Directory.Build.props"
    if (-not (Test-Path $propsPath -PathType Leaf)) {
        throw "Directory.Build.props not found: $propsPath"
    }

    try {
        [xml]$propsXml = Get-Content -Path $propsPath -Raw -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        throw "Unable to parse Directory.Build.props at $propsPath. $($_.Exception.Message)"
    }

    $versionNode = $propsXml.Project.PropertyGroup.Version | Select-Object -First 1
    $version = [string]$versionNode
    if ([string]::IsNullOrWhiteSpace($version)) {
        throw "Directory.Build.props does not define Project/PropertyGroup/Version: $propsPath"
    }

    return $version.Trim()
}

function Resolve-RepositoryPath {
    param(
        [string]$RepositoryRoot,
        [string]$PathValue
    )

    if ([string]::IsNullOrWhiteSpace($PathValue)) {
        return $null
    }

    if ([System.IO.Path]::IsPathRooted($PathValue)) {
        return [System.IO.Path]::GetFullPath($PathValue)
    }

    return [System.IO.Path]::GetFullPath((Join-Path $RepositoryRoot $PathValue))
}

function Resolve-RuntimeIdentifiers {
    param(
        [string[]]$RequestedRuntimeIdentifiers,
        [string[]]$DefaultRuntimeIdentifiers = @("win-x64")
    )

    $resolved = New-Object System.Collections.Generic.List[string]

    foreach ($runtimeIdentifier in @($RequestedRuntimeIdentifiers)) {
        if ([string]::IsNullOrWhiteSpace($runtimeIdentifier)) {
            continue
        }

        foreach ($candidateRuntimeIdentifier in @($runtimeIdentifier -split '[,;]')) {
            if ([string]::IsNullOrWhiteSpace($candidateRuntimeIdentifier)) {
                continue
            }

            $trimmed = $candidateRuntimeIdentifier.Trim()
            if (-not $resolved.Contains($trimmed)) {
                $resolved.Add($trimmed)
            }
        }
    }

    if ($resolved.Count -eq 0) {
        foreach ($runtimeIdentifier in @($DefaultRuntimeIdentifiers)) {
            if ([string]::IsNullOrWhiteSpace($runtimeIdentifier)) {
                continue
            }

            $trimmed = $runtimeIdentifier.Trim()
            if (-not $resolved.Contains($trimmed)) {
                $resolved.Add($trimmed)
            }
        }
    }

    return @($resolved.ToArray())
}

function Get-PublishArtifactsPath {
    param([string]$RepositoryRoot)

    return (Join-Path $RepositoryRoot "artifacts\\publish")
}

function Get-RuntimePublishPath {
    param(
        [string]$RepositoryRoot,
        [string]$RuntimeIdentifier
    )

    if ([string]::IsNullOrWhiteSpace($RuntimeIdentifier)) {
        throw "RuntimeIdentifier is required."
    }

    return (Join-Path (Get-PublishArtifactsPath -RepositoryRoot $RepositoryRoot) $RuntimeIdentifier.Trim())
}

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

function Assert-CommandAvailable {
    param([string]$CommandName)

    $command = Get-Command $CommandName -ErrorAction SilentlyContinue
    if ($null -eq $command) {
        throw "Required command not found: $CommandName"
    }
}

function Assert-WindowsPlatform {
    if (-not [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([System.Runtime.InteropServices.OSPlatform]::Windows)) {
        throw "This repository requires Windows."
    }
}

function Get-DotNetSdkVersion {
    Assert-CommandAvailable -CommandName "dotnet"

    $version = (& dotnet --version 2>&1 | Select-Object -Last 1).ToString().Trim()
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($version)) {
        throw "Unable to resolve the installed .NET SDK version."
    }

    return $version
}

function Get-GlobalJsonSdkSpecification {
    param([string]$RepositoryRoot)

    $globalJsonPath = Join-Path $RepositoryRoot "global.json"
    if (-not (Test-Path $globalJsonPath -PathType Leaf)) {
        throw "global.json not found: $globalJsonPath"
    }

    try {
        $globalJson = Get-Content -Path $globalJsonPath -Raw | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw "Unable to parse global.json at $globalJsonPath. $($_.Exception.Message)"
    }

    if ($null -eq $globalJson.sdk -or [string]::IsNullOrWhiteSpace($globalJson.sdk.version)) {
        throw "global.json does not define sdk.version: $globalJsonPath"
    }

    return [pscustomobject]@{
        Version = [string]$globalJson.sdk.version
        RollForward = [string]$globalJson.sdk.rollForward
        AllowPrerelease = [bool]$globalJson.sdk.allowPrerelease
        Path = $globalJsonPath
    }
}

function Assert-DotNetSdkMajorVersion {
    param([int]$MinimumMajorVersion = 8)

    $version = Get-DotNetSdkVersion
    $majorVersion = [int]($version.Split('.')[0])

    if ($majorVersion -lt $MinimumMajorVersion) {
        throw ".NET SDK $MinimumMajorVersion or later is required. Found: $version"
    }

    return $version
}

function Assert-DotNetSdkPinnedVersion {
    param([string]$RepositoryRoot)

    $sdkSpecification = Get-GlobalJsonSdkSpecification -RepositoryRoot $RepositoryRoot
    $installedVersion = Get-DotNetSdkVersion

    if (-not [string]::Equals($installedVersion, $sdkSpecification.Version, [System.StringComparison]::Ordinal)) {
        $rollForwardMessage = if ([string]::IsNullOrWhiteSpace($sdkSpecification.RollForward)) {
            "global.json pins SDK version $($sdkSpecification.Version)."
        }
        else {
            "global.json pins SDK version $($sdkSpecification.Version) with rollForward '$($sdkSpecification.RollForward)'."
        }

        throw ".NET SDK version mismatch. $rollForwardMessage Installed: $installedVersion"
    }

    return $sdkSpecification
}

function Initialize-ArtifactsLayout {
    param([string]$RepositoryRoot)

    $artifactsRoot = Join-Path $RepositoryRoot "artifacts"
    $directories = @(
        $artifactsRoot,
        (Join-Path $artifactsRoot "build"),
        (Join-Path $artifactsRoot "test-results"),
        (Join-Path $artifactsRoot "packages"),
        (Join-Path $artifactsRoot "publish"),
        (Join-Path $artifactsRoot "logs")
    )

    foreach ($directory in $directories) {
        New-Item -Path $directory -ItemType Directory -Force | Out-Null
    }

    return $artifactsRoot
}

function Invoke-DotNetCommand {
    param(
        [string[]]$Arguments,
        [string]$ErrorMessage
    )

    & dotnet @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$ErrorMessage (exit code: $LASTEXITCODE)"
    }
}

function Invoke-RepositoryPowerShellScript {
    param(
        [string]$ScriptPath,
        [string[]]$Arguments,
        [string]$ErrorMessage
    )

    if (-not (Test-Path $ScriptPath -PathType Leaf)) {
        throw "Script not found: $ScriptPath"
    }

    & pwsh -NoLogo -NoProfile -File $ScriptPath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$ErrorMessage (exit code: $LASTEXITCODE)"
    }
}

function Get-ExecutablePath {
    param(
        [string]$CommandName,
        [string[]]$CandidateDirectories = @()
    )

    foreach ($directory in $CandidateDirectories) {
        if ([string]::IsNullOrWhiteSpace($directory)) {
            continue
        }

        $candidatePath = Join-Path $directory $CommandName
        if (Test-Path $candidatePath) {
            return $candidatePath
        }
    }

    $command = Get-Command $CommandName -ErrorAction SilentlyContinue
    if ($null -ne $command -and -not [string]::IsNullOrWhiteSpace($command.Source)) {
        return $command.Source
    }

    return $null
}

function Get-InnoSetupCandidateDirectories {
    param([string]$RepositoryRoot)

    return @(
        $env:INNOSETUP_BIN,
        $(if (-not [string]::IsNullOrWhiteSpace($env:INNOSETUP_HOME)) { $env:INNOSETUP_HOME } else { $null }),
        "C:\\Program Files (x86)\\Inno Setup 6",
        "C:\\Program Files\\Inno Setup 6"
    )
}

function Get-InnoSetupBinPath {
    param([string]$RepositoryRoot)

    $candidateDirectories = Get-InnoSetupCandidateDirectories -RepositoryRoot $RepositoryRoot

    foreach ($directory in $candidateDirectories) {
        if (-not [string]::IsNullOrWhiteSpace($directory) -and (Test-Path $directory)) {
            return $directory
        }
    }

    return $null
}

function Get-InnoSetupCompilerPath {
    param([string]$RepositoryRoot)

    $candidateDirectories = Get-InnoSetupCandidateDirectories -RepositoryRoot $RepositoryRoot

    $toolPath = Get-ExecutablePath -CommandName "ISCC.exe" -CandidateDirectories $candidateDirectories
    if ($null -ne $toolPath) {
        return $toolPath
    }

    $searched = ($candidateDirectories | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) -join ", "
    throw "Inno Setup compiler not found: ISCC.exe. Configure INNOSETUP_BIN/INNOSETUP_HOME or install Inno Setup 6. Paths checked: $searched"
}

