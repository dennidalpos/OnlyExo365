#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$PublishPath,

    [string]$OutputDirectory = "artifacts/packages",

    [string]$OutputFileName = "OnlyExo365.Setup.exe",

    [string]$ProductName = "OnlyExo365",

    [string]$Manufacturer = "OnlyExo365"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$repositoryRoot = Split-Path -Parent $scriptDirectory

. (Join-Path $repositoryRoot "scripts\internal\common.ps1")

$resolvedPublishPath = Resolve-RepositoryPath -RepositoryRoot $repositoryRoot -PathValue $PublishPath
$resolvedPublishPathX64 = Join-Path $resolvedPublishPath "win-x64"
$resolvedOutputDirectory = Resolve-RepositoryPath -RepositoryRoot $repositoryRoot -PathValue $OutputDirectory
$resolvedOutputPath = Join-Path $resolvedOutputDirectory $OutputFileName
$resolvedInstallerScriptPath = Join-Path $repositoryRoot "installer\OnlyExo365.iss"
$resolvedIconPath = Join-Path $repositoryRoot "src\OnlyExo365.Shell\Assets\Generated\AppIcon.ico"
$resolvedLicensePath = Join-Path $repositoryRoot "LICENSE"
$appVersion = Get-ApplicationVersion -RepositoryRoot $repositoryRoot
$versionSegments = @($appVersion.Split('.'))
while ($versionSegments.Count -lt 4) {
    $versionSegments += "0"
}
$fileVersion = ($versionSegments | Select-Object -First 4) -join "."
$outputBaseName = [System.IO.Path]::GetFileNameWithoutExtension($OutputFileName)

if (-not (Test-Path $resolvedPublishPath -PathType Container)) {
    throw "Publish path not found: $resolvedPublishPath"
}

if (-not (Test-Path $resolvedPublishPathX64 -PathType Container)) {
    throw "Publish path not found for runtime win-x64: $resolvedPublishPathX64"
}

if (-not (Test-Path $resolvedInstallerScriptPath -PathType Leaf)) {
    throw "Installer authoring file not found: $resolvedInstallerScriptPath"
}

if (-not (Test-Path $resolvedIconPath -PathType Leaf)) {
    throw "Application icon not found: $resolvedIconPath"
}

if (-not (Test-Path $resolvedLicensePath -PathType Leaf)) {
    throw "License file not found: $resolvedLicensePath"
}

foreach ($path in @(
        (Join-Path $resolvedPublishPathX64 "OnlyExo365.Shell.exe"),
        (Join-Path $resolvedPublishPathX64 "OnlyExo365.Worker.exe"),
        (Join-Path $resolvedPublishPathX64 "appsettings.json")
    )) {
    if (-not (Test-Path $path -PathType Leaf)) {
        throw "Publish output missing required file: $path"
    }
}

$isccPath = Get-InnoSetupCompilerPath -RepositoryRoot $repositoryRoot
New-Item -Path $resolvedOutputDirectory -ItemType Directory -Force | Out-Null
Remove-Item -Path $resolvedOutputPath -Force -ErrorAction SilentlyContinue

Write-Step "Creating native setup EXE"
Write-Info "ISCC      : $isccPath"
Write-Info "Publish    : $resolvedPublishPath"
Write-Info "Publish x64: $resolvedPublishPathX64"
Write-Info "Setup EXE  : $resolvedOutputPath"
Write-Info "Version    : $appVersion"
Write-Info "FileVersion: $fileVersion"

$arguments = @(
    "/Qp",
    "/DProductName=$ProductName",
    "/DManufacturer=$Manufacturer",
    "/DAppVersion=$appVersion",
    "/DFileVersion=$fileVersion",
    "/DPublishDirX64=$resolvedPublishPathX64",
    "/DOutputDir=$resolvedOutputDirectory",
    "/DOutputBaseFilename=$outputBaseName",
    "/DIconPath=$resolvedIconPath",
    "/DLicensePath=$resolvedLicensePath",
    $resolvedInstallerScriptPath
)

& $isccPath @arguments
if ($LASTEXITCODE -ne 0) {
    throw "Inno Setup compilation failed with exit code $LASTEXITCODE."
}

if (-not (Test-Path $resolvedOutputPath -PathType Leaf)) {
    throw "Expected setup EXE not found after Inno Setup run: $resolvedOutputPath"
}

Write-Success "Setup EXE created: $resolvedOutputPath"

