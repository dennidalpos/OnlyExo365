#Requires -Version 7.0

[CmdletBinding()]
param(
    [string]$SourcePath = ".",
    [string]$ReportPath = "artifacts/security/gitleaks.sarif",
    [string]$GitleaksVersion = "8.30.0",
    [switch]$DownloadIfMissing = $true,
    [string]$ToolManifestPath = "build/tool-manifest.json"
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$repositoryRoot = Split-Path -Parent $scriptDirectory

. (Join-Path $scriptDirectory "helpers\common.ps1")

function Resolve-RepoPath {
    param(
        [string]$BaseDirectory,
        [string]$PathValue
    )

    if ([System.IO.Path]::IsPathRooted($PathValue)) {
        return $PathValue
    }

    return Join-Path $BaseDirectory $PathValue
}

function Get-FileSha256 {
    param([string]$Path)

    return (Get-FileHash -Algorithm SHA256 -Path $Path).Hash.ToLowerInvariant()
}

function Get-ToolManifest {
    param([string]$ManifestPath)

    if (-not (Test-Path $ManifestPath -PathType Leaf)) {
        Stop-WithError "Tool manifest not found: $ManifestPath"
    }

    try {
        return Get-Content -Path $ManifestPath -Raw | ConvertFrom-Json -Depth 8
    }
    catch {
        Stop-WithError "Unable to parse tool manifest $ManifestPath. $($_.Exception.Message)"
    }
}

function Get-GitleaksPackageDefinition {
    param(
        [object]$ToolManifest,
        [string]$Version
    )

    $versionProperty = $ToolManifest.gitleaks.PSObject.Properties[$Version]
    if ($null -eq $versionProperty) {
        Stop-WithError "Tool manifest does not define a pinned gitleaks package for version $Version."
    }

    $platformDefinition = $versionProperty.Value.windows_x64
    if ($null -eq $platformDefinition) {
        Stop-WithError "Tool manifest does not define a pinned windows_x64 package for gitleaks $Version."
    }

    $archiveName = [string]$platformDefinition.archiveName
    $sha256 = [string]$platformDefinition.sha256

    if ([string]::IsNullOrWhiteSpace($archiveName) -or [string]::IsNullOrWhiteSpace($sha256)) {
        Stop-WithError "Tool manifest entry for gitleaks $Version is incomplete."
    }

    if ($sha256 -notmatch '^[a-fA-F0-9]{64}$') {
        Stop-WithError "Tool manifest entry for gitleaks $Version contains an invalid SHA-256 digest."
    }

    return [pscustomobject]@{
        Version = $Version
        ArchiveName = $archiveName.Trim()
        Sha256 = $sha256.Trim().ToLowerInvariant()
    }
}

function Get-GitleaksMetadataPath {
    param([string]$VersionDirectory)

    return Join-Path $VersionDirectory "download-metadata.json"
}

function Test-GitleaksCache {
    param(
        [string]$BinaryPath,
        [string]$MetadataPath,
        [object]$ExpectedPackage
    )

    if (-not (Test-Path $BinaryPath -PathType Leaf) -or -not (Test-Path $MetadataPath -PathType Leaf)) {
        return $false
    }

    try {
        $metadata = Get-Content -Path $MetadataPath -Raw | ConvertFrom-Json -Depth 6
    }
    catch {
        Write-Info "Cached gitleaks metadata is unreadable; the tool will be re-downloaded."
        return $false
    }

    if ($metadata.version -ne $ExpectedPackage.Version -or
        $metadata.archiveName -ne $ExpectedPackage.ArchiveName -or
        $metadata.archiveSha256 -ne $ExpectedPackage.Sha256) {
        Write-Info "Cached gitleaks metadata does not match the pinned manifest; the tool will be re-downloaded."
        return $false
    }

    $actualBinarySha256 = Get-FileSha256 -Path $BinaryPath
    if ($metadata.binarySha256 -ne $actualBinarySha256) {
        Write-Info "Cached gitleaks binary hash does not match recorded metadata; the tool will be re-downloaded."
        return $false
    }

    return $true
}

function Save-GitleaksMetadata {
    param(
        [string]$MetadataPath,
        [object]$Package,
        [string]$BinaryPath
    )

    $metadata = [ordered]@{
        version = $Package.Version
        archiveName = $Package.ArchiveName
        archiveSha256 = $Package.Sha256
        binaryName = [System.IO.Path]::GetFileName($BinaryPath)
        binarySha256 = Get-FileSha256 -Path $BinaryPath
        verifiedAtUtc = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
    }

    $metadata | ConvertTo-Json -Depth 4 | Set-Content -Path $MetadataPath -Encoding UTF8
}

function Get-GitleaksExecutable {
    param(
        [string]$Version,
        [bool]$AllowDownload,
        [string]$CacheDirectory,
        [string]$ResolvedToolManifestPath
    )

    if (-not [string]::IsNullOrWhiteSpace($env:GITLEAKS_PATH) -and (Test-Path $env:GITLEAKS_PATH -PathType Leaf)) {
        return $env:GITLEAKS_PATH
    }

    $command = Get-Command gitleaks -ErrorAction SilentlyContinue
    if ($null -ne $command -and -not [string]::IsNullOrWhiteSpace($command.Source)) {
        return $command.Source
    }

    if (-not $AllowDownload) {
        return $null
    }

    $toolManifest = Get-ToolManifest -ManifestPath $ResolvedToolManifestPath
    $package = Get-GitleaksPackageDefinition -ToolManifest $toolManifest -Version $Version

    $versionDirectory = Join-Path $CacheDirectory "gitleaks-$Version"
    $binaryPath = Join-Path $versionDirectory "gitleaks.exe"
    $metadataPath = Get-GitleaksMetadataPath -VersionDirectory $versionDirectory
    if (Test-GitleaksCache -BinaryPath $binaryPath -MetadataPath $metadataPath -ExpectedPackage $package) {
        return $binaryPath
    }

    if (Test-Path $versionDirectory) {
        Remove-Item -Path $versionDirectory -Recurse -Force
    }

    New-Item -Path $versionDirectory -ItemType Directory -Force | Out-Null

    $zipPath = Join-Path $versionDirectory "gitleaks.zip"
    $downloadUrl = "https://github.com/gitleaks/gitleaks/releases/download/v$Version/$($package.ArchiveName)"

    Write-Info "Downloading gitleaks v$Version from $downloadUrl"
    Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath

    $actualArchiveSha256 = Get-FileSha256 -Path $zipPath
    if ($actualArchiveSha256 -ne $package.Sha256) {
        Stop-WithError "Downloaded gitleaks archive hash mismatch. Expected $($package.Sha256), actual $actualArchiveSha256."
    }

    Expand-Archive -Path $zipPath -DestinationPath $versionDirectory -Force
    Remove-Item -Path $zipPath -Force

    if (-not (Test-Path $binaryPath -PathType Leaf)) {
        Stop-WithError "Downloaded gitleaks archive does not contain gitleaks.exe."
    }

    Save-GitleaksMetadata -MetadataPath $metadataPath -Package $package -BinaryPath $binaryPath

    return $binaryPath
}

$resolvedSourcePath = Resolve-RepoPath -BaseDirectory $repositoryRoot -PathValue $SourcePath
$resolvedReportPath = Resolve-RepoPath -BaseDirectory $repositoryRoot -PathValue $ReportPath
$resolvedToolManifestPath = Resolve-RepoPath -BaseDirectory $repositoryRoot -PathValue $ToolManifestPath
$reportDirectory = Split-Path -Parent $resolvedReportPath
$cacheDirectory = Join-Path ([System.IO.Path]::GetTempPath()) "onlyexo365-security-tools"

if (-not (Test-Path $resolvedSourcePath)) {
    Stop-WithError "Source path not found: $resolvedSourcePath"
}

if (-not (Test-Path $reportDirectory)) {
    New-Item -Path $reportDirectory -ItemType Directory -Force | Out-Null
}

if (-not (Test-Path $cacheDirectory)) {
    New-Item -Path $cacheDirectory -ItemType Directory -Force | Out-Null
}

$gitleaksPath = Get-GitleaksExecutable -Version $GitleaksVersion -AllowDownload:$DownloadIfMissing.IsPresent -CacheDirectory $cacheDirectory -ResolvedToolManifestPath $resolvedToolManifestPath
if ([string]::IsNullOrWhiteSpace($gitleaksPath)) {
    Stop-WithError "gitleaks executable not found and download disabled."
}

Write-Step "Scanning repository for committed secrets"
Write-Info "Source  : $resolvedSourcePath"
Write-Info "Report  : $resolvedReportPath"
Write-Info "gitleaks: $gitleaksPath"

& $gitleaksPath dir $resolvedSourcePath --redact --report-format sarif --report-path $resolvedReportPath --exit-code 1
$exitCode = $LASTEXITCODE

if ($exitCode -eq 0) {
    Write-Success "No secrets detected by gitleaks."
    exit 0
}

if ($exitCode -eq 1) {
    Stop-WithError "Secret scan found one or more findings. See $resolvedReportPath." 1
}

Stop-WithError "gitleaks failed unexpectedly (exit code: $exitCode)." $exitCode
