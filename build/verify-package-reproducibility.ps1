#Requires -Version 7.0

[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",

    [ValidateNotNullOrEmpty()]
    [string]$RuntimeIdentifier = "win-x64",

    [string]$ReportPath = "artifacts/reproducibility/package-reproducibility.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$repositoryRoot = Split-Path -Parent $scriptDirectory
$packScriptPath = Join-Path $repositoryRoot "scripts\pack.ps1"
$artifactsRoot = Join-Path $repositoryRoot "artifacts"
$publishPath = Join-Path $artifactsRoot "publish"
$packagesPath = Join-Path $artifactsRoot "packages"
$resolvedReportPath = if ([System.IO.Path]::IsPathRooted($ReportPath)) {
    [System.IO.Path]::GetFullPath($ReportPath)
}
else {
    [System.IO.Path]::GetFullPath((Join-Path $repositoryRoot $ReportPath))
}
$snapshotRoot = Join-Path $artifactsRoot "reproducibility\snapshots"

. (Join-Path $repositoryRoot "scripts\internal\common.ps1")

function Get-ArtifactHashMap {
    param([string]$RootPath)

    $items = [ordered]@{}
    foreach ($file in Get-ChildItem -Path $RootPath -Recurse -File | Sort-Object FullName) {
        $relativePath = [System.IO.Path]::GetRelativePath($RootPath, $file.FullName).Replace('\', '/')
        $items[$relativePath] = (Get-FileHash -Algorithm SHA256 -Path $file.FullName).Hash.ToLowerInvariant()
    }

    return $items
}

function Copy-Snapshot {
    param(
        [string]$SourcePath,
        [string]$DestinationPath
    )

    if (Test-Path $DestinationPath) {
        Remove-Item -Path $DestinationPath -Recurse -Force
    }

    New-Item -Path $DestinationPath -ItemType Directory -Force | Out-Null
    Copy-Item -Path (Join-Path $SourcePath '*') -Destination $DestinationPath -Recurse -Force
}

function Invoke-PackRun {
    param([int]$Iteration)

    Write-Step "Running pack iteration $Iteration"
    & pwsh -NoLogo -NoProfile -File $packScriptPath `
        -Configuration $Configuration `
        -Clean:$true `
        -LockedMode `
        -RuntimeIdentifier $RuntimeIdentifier | Out-Host

    if ($LASTEXITCODE -ne 0) {
        throw "scripts/pack.ps1 failed during iteration $Iteration with exit code $LASTEXITCODE."
    }

    $publishSnapshotPath = Join-Path $snapshotRoot "run-$Iteration\publish"
    $packagesSnapshotPath = Join-Path $snapshotRoot "run-$Iteration\packages"
    Copy-Snapshot -SourcePath $publishPath -DestinationPath $publishSnapshotPath
    Copy-Snapshot -SourcePath $packagesPath -DestinationPath $packagesSnapshotPath

    return [pscustomobject]@{
        PublishPath = $publishSnapshotPath
        PackagesPath = $packagesSnapshotPath
        PublishHashes = Get-ArtifactHashMap -RootPath $publishSnapshotPath
        PackageHashes = Get-ArtifactHashMap -RootPath $packagesSnapshotPath
    }
}

New-Item -Path (Split-Path -Parent $resolvedReportPath) -ItemType Directory -Force | Out-Null
if (Test-Path $snapshotRoot) {
    Remove-Item -Path $snapshotRoot -Recurse -Force
}

$firstRun = Invoke-PackRun -Iteration 1
$secondRun = Invoke-PackRun -Iteration 2

$mismatches = @()

foreach ($category in @(
        @{ Name = "publish"; First = $firstRun.PublishHashes; Second = $secondRun.PublishHashes },
        @{ Name = "packages"; First = $firstRun.PackageHashes; Second = $secondRun.PackageHashes }
    )) {
    $allKeys = @($category.First.Keys + $category.Second.Keys | Sort-Object -Unique)
    foreach ($key in $allKeys) {
        $firstHash = if ($category.First.Contains($key)) { $category.First[$key] } else { $null }
        $secondHash = if ($category.Second.Contains($key)) { $category.Second[$key] } else { $null }
        if ($firstHash -ne $secondHash) {
            $mismatches += [pscustomobject]@{
                category = $category.Name
                path = $key
                first_hash = $firstHash
                second_hash = $secondHash
            }
        }
    }
}

$report = [ordered]@{
    executed_on = (Get-Date).ToString("yyyy-MM-dd")
    configuration = $Configuration
    runtime_identifier = $RuntimeIdentifier
    publish_reproducible = @($mismatches | Where-Object category -eq "publish").Count -eq 0
    package_reproducible = @($mismatches | Where-Object category -eq "packages").Count -eq 0
    mismatches = $mismatches
}

$report | ConvertTo-Json -Depth 6 | Set-Content -Path $resolvedReportPath -Encoding UTF8

if ($mismatches.Count -gt 0) {
    throw "Package reproducibility check failed. See report: $resolvedReportPath"
}

Write-Success "Package reproducibility verified. Report: $resolvedReportPath"
