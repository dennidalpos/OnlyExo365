#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$PublishPath,

    [Parameter(Mandatory = $true)]
    [string]$SetupExePath,

    [Parameter(Mandatory = $true)]
    [string]$ReleaseTag,

    [ValidateNotNullOrEmpty()]
    [string]$RuntimeIdentifier = "win-x64",

    [string]$OutputDirectory = "artifacts/publish/release-assets"
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$repositoryRoot = Split-Path -Parent $scriptDirectory

. (Join-Path $scriptDirectory "helpers\common.ps1")

function Resolve-AbsolutePath {
    param([string]$PathValue)

    if ([string]::IsNullOrWhiteSpace($PathValue)) {
        return $null
    }

    try {
        return (Resolve-Path -Path $PathValue -ErrorAction Stop).Path
    }
    catch {
        if ([System.IO.Path]::IsPathRooted($PathValue)) {
            return [System.IO.Path]::GetFullPath($PathValue)
        }

        return [System.IO.Path]::GetFullPath((Join-Path $repositoryRoot $PathValue))
    }
}

function Get-ReleaseVersionFromExecutable {
    param([string]$Path)

    $file = Get-Item -Path $Path -ErrorAction Stop
    $version = $file.VersionInfo.ProductVersion
    if ([string]::IsNullOrWhiteSpace($version)) {
        return "0.0.0"
    }

    return $version.Split('+')[0].Trim()
}

function Get-FileSha256 {
    param([string]$Path)

    return (Get-FileHash -Algorithm SHA256 -Path $Path).Hash.ToLowerInvariant()
}

$resolvedPublishPath = Resolve-AbsolutePath -PathValue $PublishPath
$resolvedSetupExePath = Resolve-AbsolutePath -PathValue $SetupExePath
$resolvedOutputDirectory = Resolve-AbsolutePath -PathValue $OutputDirectory
$resolvedRuntimeIdentifier = $RuntimeIdentifier.Trim()

if (-not (Test-Path $resolvedPublishPath -PathType Container)) {
    Stop-WithError "Publish path not found or not a directory: $PublishPath"
}

if (-not (Test-Path $resolvedSetupExePath -PathType Leaf)) {
    Stop-WithError "Setup EXE path not found: $SetupExePath"
}

if ([string]::IsNullOrWhiteSpace($ReleaseTag)) {
    Stop-WithError "ReleaseTag is required."
}

if ([string]::IsNullOrWhiteSpace($resolvedRuntimeIdentifier)) {
    Stop-WithError "RuntimeIdentifier is required."
}

$sanitizedTag = ($ReleaseTag.Trim() -replace '[^A-Za-z0-9._-]', '-')
$releaseVersion = Get-ReleaseVersionFromExecutable -Path $resolvedSetupExePath
$zipFileName = "OnlyExo365-$sanitizedTag-$resolvedRuntimeIdentifier-publish.zip"
$setupExeFileName = "OnlyExo365-$sanitizedTag-$resolvedRuntimeIdentifier-setup.exe"
$checksumsFileName = "OnlyExo365-$sanitizedTag-$resolvedRuntimeIdentifier.sha256"
$manifestFileName = "OnlyExo365-$sanitizedTag-$resolvedRuntimeIdentifier-assets.json"
$zipDestination = Join-Path $resolvedOutputDirectory $zipFileName
$setupExeDestination = Join-Path $resolvedOutputDirectory $setupExeFileName
$checksumsPath = Join-Path $resolvedOutputDirectory $checksumsFileName
$manifestPath = Join-Path $resolvedOutputDirectory $manifestFileName
$temporaryZipPath = Join-Path ([System.IO.Path]::GetTempPath()) ("onlyexo365-release-assets-" + [guid]::NewGuid().ToString("N") + ".zip")

Write-Step "Preparing release assets"
Write-Info "PublishPath : $resolvedPublishPath"
Write-Info "SetupExePath: $resolvedSetupExePath"
Write-Info "ReleaseTag  : $ReleaseTag"
Write-Info "Runtime     : $resolvedRuntimeIdentifier"
Write-Info "OutputDir   : $resolvedOutputDirectory"

New-Item -Path $resolvedOutputDirectory -ItemType Directory -Force | Out-Null
Remove-Item -Path $zipDestination, $setupExeDestination, $checksumsPath, $manifestPath -Force -ErrorAction SilentlyContinue

$publishArchiveInputs = @(Get-ChildItem -Path $resolvedPublishPath -Force | Where-Object {
        -not [string]::Equals($_.FullName, $resolvedOutputDirectory, [System.StringComparison]::OrdinalIgnoreCase)
    } | Sort-Object FullName | Select-Object -ExpandProperty FullName)
$publishFiles = @(Get-ChildItem -Path $resolvedPublishPath -Recurse -File | Where-Object {
        -not $_.FullName.StartsWith($resolvedOutputDirectory, [System.StringComparison]::OrdinalIgnoreCase)
    } | Sort-Object FullName)
if ($publishFiles.Count -eq 0) {
    Stop-WithError "Publish directory is empty: $resolvedPublishPath"
}

if ($publishArchiveInputs.Count -eq 0) {
    Stop-WithError "Publish directory does not contain any files outside the release assets output directory: $resolvedPublishPath"
}

Remove-Item -Path $temporaryZipPath -Force -ErrorAction SilentlyContinue
Compress-Archive -Path $publishArchiveInputs -DestinationPath $temporaryZipPath -CompressionLevel Optimal
Move-Item -Path $temporaryZipPath -Destination $zipDestination -Force
Copy-Item -Path $resolvedSetupExePath -Destination $setupExeDestination -Force

$zipSha256 = Get-FileSha256 -Path $zipDestination
$setupExeSha256 = Get-FileSha256 -Path $setupExeDestination

@(
    "$zipSha256 *$zipFileName"
    "$setupExeSha256 *$setupExeFileName"
) | Set-Content -Path $checksumsPath -Encoding UTF8

$manifest = [ordered]@{
    release_tag = $ReleaseTag.Trim()
    version = $releaseVersion
    runtime_identifier = $resolvedRuntimeIdentifier
    publish_zip = [ordered]@{
        file_name = $zipFileName
        sha256 = $zipSha256
        file_count = $publishFiles.Count
    }
    setup_exe = [ordered]@{
        file_name = $setupExeFileName
        sha256 = $setupExeSha256
    }
    checksums = [ordered]@{
        file_name = $checksumsFileName
    }
}

$manifest | ConvertTo-Json -Depth 6 | Set-Content -Path $manifestPath -Encoding UTF8

Write-Success "Prepared release publish zip: $zipDestination"
Write-Success "Prepared release setup EXE: $setupExeDestination"
Write-Success "Prepared checksums        : $checksumsPath"
Write-Success "Prepared manifest         : $manifestPath"

if (-not [string]::IsNullOrWhiteSpace($env:GITHUB_ENV)) {
    "RELEASE_ASSET_ZIP=$zipDestination" | Out-File -FilePath $env:GITHUB_ENV -Encoding utf8 -Append
    "RELEASE_ASSET_SETUP_EXE=$setupExeDestination" | Out-File -FilePath $env:GITHUB_ENV -Encoding utf8 -Append
    "RELEASE_ASSET_CHECKSUMS=$checksumsPath" | Out-File -FilePath $env:GITHUB_ENV -Encoding utf8 -Append
    "RELEASE_ASSET_MANIFEST=$manifestPath" | Out-File -FilePath $env:GITHUB_ENV -Encoding utf8 -Append
}

Remove-Item -Path $temporaryZipPath -Force -ErrorAction SilentlyContinue
