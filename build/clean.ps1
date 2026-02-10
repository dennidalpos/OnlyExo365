#Requires -Version 7.0

[CmdletBinding()]
param(
    [switch]$All,
    [switch]$DryRun,
    [switch]$SkipDotNetClean,
    [switch]$IncludeExports
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$SolutionDir = Split-Path -Parent $ScriptDir
$ArtifactsDir = Join-Path $SolutionDir "artifacts"
$SrcDir = Join-Path $SolutionDir "src"
$ExportsDir = Join-Path $ArtifactsDir "exports"

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

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host ">> $Message" -ForegroundColor Cyan
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

function Write-Success {
    param([string]$Message)
    Write-Host "   [OK] $Message" -ForegroundColor Green
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

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " ExchangeAdmin Clean Script" -ForegroundColor Cyan
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
if (Remove-DirectoryIfExists -Path $ArtifactsDir -Description "Artifacts") {
    Write-Success "Artifacts directory cleaned"
} else {
    Write-Skipped "Artifacts directory not found"
}


Write-Step "Cleaning generated export files"
if ($IncludeExports) {
    if (Remove-DirectoryIfExists -Path $ExportsDir -Description "Generated exports") {
        Write-Success "Generated exports cleaned"
    } else {
        Write-Skipped "Generated exports directory not found"
    }
} else {
    Write-Skipped "Generated exports clean skipped (use -IncludeExports)"
}

Write-Step "Cleaning bin/obj directories"

# Clean top-level bin/obj folders if present (for uncommon solution-level outputs).
$rootBinObj = @(
    (Join-Path $SolutionDir 'bin'),
    (Join-Path $SolutionDir 'obj')
)
foreach ($dir in $rootBinObj) {
    [void](Remove-DirectoryIfExists -Path $dir -Description 'Solution-level build output')
}

$projects = Get-ChildItem -Path $SrcDir -Directory -ErrorAction SilentlyContinue
$cleanedProjects = 0

foreach ($project in $projects) {
    $binDir = Join-Path $project.FullName "bin"
    $objDir = Join-Path $project.FullName "obj"

    $binCleaned = Remove-DirectoryIfExists -Path $binDir -Description "$($project.Name)/bin"
    $objCleaned = Remove-DirectoryIfExists -Path $objDir -Description "$($project.Name)/obj"

    if ($binCleaned -or $objCleaned) {
        $cleanedProjects++
    }
}

if ($cleanedProjects -gt 0) {
    Write-Success "Cleaned $cleanedProjects project(s)"
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
$solutionFile = Join-Path $SolutionDir "ExchangeAdmin.sln"

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
