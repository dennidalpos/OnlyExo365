[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$CoveragePath,

    [Parameter(Mandatory = $true)]
    [double]$MinimumLineCoveragePercent,

    [string[]]$IncludePackage = @(),

    [string[]]$ExcludeFilePattern = @('\\obj\\')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$normalizedPackages = @()
foreach ($package in $IncludePackage) {
    if ([string]::IsNullOrWhiteSpace($package)) {
        continue
    }

    $normalizedPackages += $package.Split(',', [System.StringSplitOptions]::RemoveEmptyEntries -bor [System.StringSplitOptions]::TrimEntries)
}

if (-not (Test-Path -LiteralPath $CoveragePath)) {
    throw "Coverage file not found: $CoveragePath"
}

[xml]$coverageDocument = Get-Content -LiteralPath $CoveragePath
$packageNodes = @($coverageDocument.coverage.packages.package)

if ($normalizedPackages.Count -gt 0) {
    $packageNodes = @($packageNodes | Where-Object { $_.name -in $normalizedPackages })
}

if ($packageNodes.Count -eq 0) {
    throw "No coverage packages matched the requested filter."
}

$packageResults = New-Object System.Collections.Generic.List[object]
$totalLines = 0
$coveredLines = 0

foreach ($packageNode in $packageNodes) {
    $packageTotal = 0
    $packageCovered = 0

    foreach ($classNode in @($packageNode.classes.class)) {
        $filePath = [string]$classNode.filename
        if ($ExcludeFilePattern | Where-Object { $filePath -match $_ }) {
            continue
        }

        $lineNodes = @($classNode.SelectNodes('lines/line'))
        if ($lineNodes.Count -eq 0) {
            continue
        }

        foreach ($lineNode in $lineNodes) {
            $packageTotal++
            if ([int]$lineNode.hits -gt 0) {
                $packageCovered++
            }
        }
    }

    if ($packageTotal -eq 0) {
        continue
    }

    $packagePercent = [math]::Round(($packageCovered / [double]$packageTotal) * 100, 2)
    $packageResults.Add([pscustomobject]@{
        Package = [string]$packageNode.name
        CoveredLines = $packageCovered
        TotalLines = $packageTotal
        LineCoveragePercent = $packagePercent
    })

    $coveredLines += $packageCovered
    $totalLines += $packageTotal
}

if ($totalLines -eq 0) {
    throw "Coverage report did not contain any executable lines after filters were applied."
}

$overallPercent = [math]::Round(($coveredLines / [double]$totalLines) * 100, 2)

Write-Host "Coverage summary from $CoveragePath"
$packageResults | Sort-Object Package | Format-Table -AutoSize | Out-String | Write-Host
Write-Host ("Overall line coverage: {0:N2}% ({1}/{2})" -f $overallPercent, $coveredLines, $totalLines)

if ($overallPercent -lt $MinimumLineCoveragePercent) {
    throw ("Line coverage gate failed. Expected at least {0:N2}%, actual {1:N2}%." -f $MinimumLineCoveragePercent, $overallPercent)
}
