#Requires -Version 7.0

[CmdletBinding()]
param(
    [switch]$All,
    [switch]$DryRun,
    [switch]$SkipDotNetClean,
    [string]$ExportDirPath,
    [string]$ImportDirPath
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$SolutionDir = Split-Path -Parent $ScriptDir
$ArtifactsDir = Join-Path $SolutionDir "artifacts"
$BuildArtifactsDir = Join-Path $ArtifactsDir "build"
$ExportsDir = if ([string]::IsNullOrWhiteSpace($ExportDirPath)) { Join-Path $ArtifactsDir "exports" } else { $ExportDirPath }
$ImportsDir = if ([string]::IsNullOrWhiteSpace($ImportDirPath)) { Join-Path $ArtifactsDir "imports" } else { $ImportDirPath }

. (Join-Path $ScriptDir "helpers\common.ps1")

if (-not [System.IO.Path]::IsPathRooted($ExportsDir)) {
    $ExportsDir = Join-Path $SolutionDir $ExportsDir
}
if (-not [System.IO.Path]::IsPathRooted($ImportsDir)) {
    $ImportsDir = Join-Path $SolutionDir $ImportsDir
}

$script:DeletedCount = 0
$script:DeletedSize = 0

function Test-DotNetAvailable {
    try {
        $null = & dotnet --version 2>&1
        return $LASTEXITCODE -eq 0
    }
    catch {
        return $false
    }
}

function Write-Deleted {
    param([string]$Path, [long]$Size = 0)
    $sizeStr = if ($Size -gt 0) { " ({0:N2} MB)" -f ($Size / 1MB) } else { "" }
    if ($DryRun) {
        Write-Host "   [DRY-RUN] Would delete: $Path$sizeStr" -ForegroundColor Yellow
    } else {
        Write-Host "   [DEL] $Path$sizeStr" -ForegroundColor Red
    }
    $script:DeletedCount++
    $script:DeletedSize += $Size
}

function Write-Skipped {
    param([string]$Message)
    Write-Host "   [SKIP] $Message" -ForegroundColor Gray
}

function Get-DirectorySize {
    param([string]$Path)
    if (Test-Path $Path) {
        return (Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue |
                Measure-Object -Property Length -Sum).Sum
    }
    return 0
}

function Remove-DirectoryIfExists {
    param([string]$Path, [string]$Description)

    if (Test-Path $Path) {
        $size = Get-DirectorySize -Path $Path
        Write-Deleted -Path $Path -Size $size

        if (-not $DryRun) {
            Remove-Item -Path $Path -Recurse -Force -ErrorAction SilentlyContinue
        }
        return $true
    }
    return $false
}

function Remove-FilesByPattern {
    param(
        [string]$Path,
        [string[]]$Patterns,
        [string]$Description
    )

    $found = $false
    foreach ($pattern in $Patterns) {
        $files = Get-ChildItem -Path $Path -Filter $pattern -Recurse -File -ErrorAction SilentlyContinue
        foreach ($file in $files) {
            # Never touch Git internals.
            if ($file.FullName -like '*\.git\*' -or $file.FullName -like '*/.git/*') {
                continue
            }

            Write-Deleted -Path $file.FullName -Size $file.Length
            $found = $true

            if (-not $DryRun) {
                Remove-Item -Path $file.FullName -Force -ErrorAction SilentlyContinue
            }
        }
    }
    return $found
}

function Get-RepoRelativePath {
    param([string]$Path)

    try {
        return [System.IO.Path]::GetRelativePath($SolutionDir, $Path).Replace('\', '/')
    }
    catch {
        return $Path
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " OnlyExo365 Clean Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Solution: $SolutionDir"
if ($DryRun) {
    Write-Host "Mode: DRY-RUN (nothing will be deleted)" -ForegroundColor Yellow
}
if ($All) {
    Write-Host "Scope: ALL (including caches)" -ForegroundColor Yellow
}


Write-Step "Cleaning artifacts directory"
if (Test-Path $ArtifactsDir) {
    $artifactEntries = Get-ChildItem -Path $ArtifactsDir -Force -ErrorAction SilentlyContinue

    foreach ($entry in $artifactEntries) {
        if ($entry.PSIsContainer) {
            [void](Remove-DirectoryIfExists -Path $entry.FullName -Description "artifacts/$($entry.Name)")
        } else {
            Write-Deleted -Path $entry.FullName -Size $entry.Length
            if (-not $DryRun) {
                Remove-Item -Path $entry.FullName -Force -ErrorAction SilentlyContinue
            }
        }
    }

    Write-Success "Artifacts directory cleaned"
} else {
    Write-Skipped "Artifacts directory not found"
}

if ($PSBoundParameters.ContainsKey('ExportDirPath')) {
    Write-Step "Cleaning custom export directory"
    if (Remove-DirectoryIfExists -Path $ExportsDir -Description "Generated exports") {
        Write-Success "Custom export directory cleaned"
    } else {
        Write-Skipped "Custom export directory not found"
    }
}

if ($PSBoundParameters.ContainsKey('ImportDirPath')) {
    Write-Step "Cleaning custom import directory"
    if (Remove-DirectoryIfExists -Path $ImportsDir -Description "Generated imports") {
        Write-Success "Custom import directory cleaned"
    } else {
        Write-Skipped "Custom import directory not found"
    }
}

Write-Step "Cleaning bin/obj directories"
$binObjDirectories = @(Get-ChildItem -Path $SolutionDir -Directory -Recurse -Force -ErrorAction SilentlyContinue |
    Where-Object {
        $_.Name -in @('bin', 'obj') -and
        $_.FullName -notlike "$ArtifactsDir\*" -and
        $_.FullName -notlike "$SolutionDir\.git\*"
    } |
    Sort-Object FullName -Unique)

$cleanedBinObjDirectories = 0

foreach ($directory in $binObjDirectories) {
    $description = Get-RepoRelativePath -Path $directory.FullName
    if (Remove-DirectoryIfExists -Path $directory.FullName -Description $description) {
        $cleanedBinObjDirectories++
    }
}

if ($cleanedBinObjDirectories -gt 0) {
    Write-Success "Cleaned $cleanedBinObjDirectories bin/obj directorie(s)"
} else {
    Write-Skipped "No bin/obj directories found"
}

Write-Step "Cleaning temporary files"
$tempPatterns = @(
    "*.tmp",
    "*.temp",
    "*.log",
    "*.bak",
    "*.orig",
    "*~",
    "*.cache"
)

$tempCleaned = Remove-FilesByPattern -Path $SolutionDir -Patterns $tempPatterns -Description "Temporary files"
if ($tempCleaned) {
    Write-Success "Temporary files cleaned"
} else {
    Write-Skipped "No temporary files found"
}

Write-Step "Cleaning test results and coverage"
$testDirs = @(
    (Join-Path $SolutionDir "TestResults"),
    (Join-Path $SolutionDir "coverage"),
    (Join-Path $SolutionDir ".coverage")
)

$recursiveTestDirectories = @(Get-ChildItem -Path $SolutionDir -Directory -Recurse -Force -ErrorAction SilentlyContinue |
    Where-Object {
        $_.Name -in @('TestResults', 'coverage', '.coverage') -and
        $_.FullName -notlike "$ArtifactsDir\*" -and
        $_.FullName -notlike "$SolutionDir\.git\*"
    } |
    Sort-Object FullName -Unique |
    Select-Object -ExpandProperty FullName)

foreach ($directory in $recursiveTestDirectories) {
    if ($directory -notin $testDirs) {
        $testDirs += $directory
    }
}

$testCleaned = $false
foreach ($testDir in $testDirs) {
    if (Remove-DirectoryIfExists -Path $testDir -Description "Test results") {
        $testCleaned = $true
    }
}

$testPatterns = @(
    "*.trx",
    "coverage.cobertura.xml",
    "coverage.opencover.xml",
    "*.coverage"
)

if (Remove-FilesByPattern -Path $SolutionDir -Patterns $testPatterns -Description "Test coverage files") {
    $testCleaned = $true
}

if ($testCleaned) {
    Write-Success "Test results cleaned"
} else {
    Write-Skipped "No test results found"
}

Write-Step "Cleaning IDE temporary files"
$ideDirs = @(
    (Join-Path $SolutionDir ".vs"),
    (Join-Path $SolutionDir '.idea')
)

$ideCleaned = $false
foreach ($ideDir in $ideDirs) {
    if (Test-Path $ideDir) {
        if (Remove-DirectoryIfExists -Path $ideDir -Description "IDE cache") {
            $ideCleaned = $true
        }
    }
}

$idePatterns = @(
    "*.suo",
    "*.user",
    "*.DotSettings.user"
)

if (Remove-FilesByPattern -Path $SolutionDir -Patterns $idePatterns -Description "IDE user files") {
    $ideCleaned = $true
}

if ($ideCleaned) {
    Write-Success "IDE files cleaned"
} else {
    Write-Skipped "No IDE temporary files found"
}

if ($All) {
    Write-Step "Cleaning NuGet caches"

    $packagesDir = Join-Path $SolutionDir "packages"
    if (Remove-DirectoryIfExists -Path $packagesDir -Description "Local packages") {
        Write-Success "Local NuGet packages cleaned"
    } else {
        Write-Skipped "No local packages directory"
    }

    if (-not $DryRun) {
        if (Test-DotNetAvailable) {
            Write-Host "   Clearing NuGet HTTP cache..." -ForegroundColor Gray
            & dotnet nuget locals http-cache --clear 2>&1 | Out-Null
            Write-Success "NuGet HTTP cache cleared"
        } else {
            Write-Skipped "dotnet SDK not available; skipped NuGet cache clean"
        }
    } else {
        Write-Host "   [DRY-RUN] Would clear NuGet HTTP cache" -ForegroundColor Yellow
    }
}

Write-Step "Running dotnet clean"
$solutionFile = Join-Path $SolutionDir "OnlyExo365.sln"

if ($SkipDotNetClean) {
    Write-Skipped 'dotnet clean skipped by -SkipDotNetClean'
}
elseif (Test-Path $solutionFile) {
    if (-not $DryRun) {
        if (Test-DotNetAvailable) {
            & dotnet clean $solutionFile --verbosity minimal 2>&1 | ForEach-Object {
                if ($_ -match 'error') {
                    Write-Host "   $_" -ForegroundColor Red
                }
            }
            if (Test-Path $BuildArtifactsDir) {
                & dotnet clean $solutionFile --artifacts-path $BuildArtifactsDir --verbosity minimal 2>&1 | ForEach-Object {
                    if ($_ -match 'error') {
                        Write-Host "   $_" -ForegroundColor Red
                    }
                }
            }
            Write-Success "dotnet clean completed"
        } else {
            Write-Skipped "dotnet SDK not available; skipped dotnet clean"
        }
    } else {
        Write-Host "   [DRY-RUN] Would run: dotnet clean $solutionFile" -ForegroundColor Yellow
    }
} else {
    Write-Skipped "Solution file not found"
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " CLEAN COMPLETED" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

if ($DryRun) {
    Write-Host "DRY-RUN Summary:" -ForegroundColor Yellow
    Write-Host "  Items that would be deleted: $script:DeletedCount"
    Write-Host ("  Space that would be freed: {0:N2} MB" -f ($script:DeletedSize / 1MB))
    Write-Host ""
    Write-Host "Run without -DryRun to actually delete files." -ForegroundColor Yellow
} else {
    Write-Host "Cleaned $script:DeletedCount item(s)"
    Write-Host ("Freed {0:N2} MB of disk space" -f ($script:DeletedSize / 1MB))
}

Write-Host ""
exit 0

