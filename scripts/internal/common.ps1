#Requires -Version 7.0

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-RepositoryRoot {
    param([string]$ScriptRoot)

    if ([string]::IsNullOrWhiteSpace($ScriptRoot)) {
        throw "ScriptRoot is required."
    }

    $currentDirectory = [System.IO.DirectoryInfo]::new([System.IO.Path]::GetFullPath($ScriptRoot))
    while ($null -ne $currentDirectory) {
        $solutionPath = Join-Path $currentDirectory.FullName "OnlyExo365.sln"
        $globalJsonPath = Join-Path $currentDirectory.FullName "global.json"
        if ((Test-Path $solutionPath -PathType Leaf) -and (Test-Path $globalJsonPath -PathType Leaf)) {
            return $currentDirectory.FullName
        }

        $currentDirectory = $currentDirectory.Parent
    }

    throw "Repository root not found from script root: $ScriptRoot"
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

function Get-DotNetInstalledSdkVersions {
    Assert-CommandAvailable -CommandName "dotnet"

    $output = @(& dotnet --list-sdks 2>&1)
    if ($LASTEXITCODE -ne 0) {
        $details = ($output | ForEach-Object { $_.ToString().Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join " "
        if ([string]::IsNullOrWhiteSpace($details)) {
            throw "Unable to list installed .NET SDK versions."
        }

        throw "Unable to list installed .NET SDK versions. $details"
    }

    $versions = @(
        $output |
            ForEach-Object {
                $line = $_.ToString().Trim()
                if ($line -match '^(?<version>\d+\.\d+\.\d+(?:[-a-zA-Z0-9.]+)?)\s+\[') {
                    $Matches.version
                }
            } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    return @($versions | Sort-Object -Unique)
}

function Assert-WindowsPlatform {
    if (-not [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([System.Runtime.InteropServices.OSPlatform]::Windows)) {
        throw "This repository requires Windows."
    }
}

function Get-DotNetSdkVersion {
    Assert-CommandAvailable -CommandName "dotnet"

    $output = @(& dotnet --version 2>&1)
    $lastOutput = $output | Select-Object -Last 1
    $version = if ($null -eq $lastOutput) { "" } else { $lastOutput.ToString().Trim() }
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($version)) {
        $installedVersions = @(Get-DotNetInstalledSdkVersions)
        $details = ($output | ForEach-Object { $_.ToString().Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join " "
        $installedMessage = if ($installedVersions.Count -gt 0) { " Installed SDKs: $($installedVersions -join ', ')." } else { "" }
        if ([string]::IsNullOrWhiteSpace($details)) {
            throw "Unable to resolve the active .NET SDK version.$installedMessage"
        }

        throw "Unable to resolve the active .NET SDK version. $details$installedMessage"
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
    $installedVersions = @(Get-DotNetInstalledSdkVersions)

    if ($installedVersions -notcontains $sdkSpecification.Version) {
        $rollForwardMessage = if ([string]::IsNullOrWhiteSpace($sdkSpecification.RollForward)) {
            "global.json pins SDK version $($sdkSpecification.Version)."
        }
        else {
            "global.json pins SDK version $($sdkSpecification.Version) with rollForward '$($sdkSpecification.RollForward)'."
        }

        $installedMessage = if ($installedVersions.Count -gt 0) {
            "Installed SDKs: $($installedVersions -join ', ')."
        }
        else {
            "No .NET SDK versions were reported by 'dotnet --list-sdks'."
        }

        throw ".NET SDK version mismatch. $rollForwardMessage $installedMessage Install SDK $($sdkSpecification.Version) or update global.json to an installed SDK."
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

function Invoke-RepositoryBootstrap {
    param(
        [string]$RepositoryRoot,
        [bool]$LockedMode,
        [string[]]$RuntimeIdentifiers,
        [switch]$Skip
    )

    if ($Skip) {
        return
    }

    $resolvedRuntimeIdentifiers = Resolve-RuntimeIdentifiers -RequestedRuntimeIdentifiers $RuntimeIdentifiers -DefaultRuntimeIdentifiers @("win-x64")
    Invoke-RepositoryPowerShellScript `
        -ScriptPath (Join-Path $RepositoryRoot "scripts/bootstrap.ps1") `
        -Arguments @("-LockedMode:$LockedMode", "-RuntimeIdentifiers", ($resolvedRuntimeIdentifiers -join ',')) `
        -ErrorMessage "bootstrap failed"
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

