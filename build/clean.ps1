#Requires -Version 7.0
[CmdletBinding()]
param(
    [switch]$All,
    [switch]$DryRun,
    [switch]$SkipDotNetClean,
    [string]$ExportDirPath,
    [string]$ImportDirPath
)

$SolutionDir = Split-Path -Parent $PSScriptRoot
. (Join-Path $SolutionDir 'scripts\internal\common.ps1')
$targets = [Collections.Generic.List[string]]::new()
foreach ($target in @(Get-RepositoryCleanupTargets -RepositoryRoot $SolutionDir)) { $targets.Add($target) }

foreach ($option in @(@{ Name='ExportDirPath'; Value=$ExportDirPath; Root='exports' }, @{ Name='ImportDirPath'; Value=$ImportDirPath; Root='imports' })) {
    if (-not $PSBoundParameters.ContainsKey($option.Name)) { continue }
    $resolved = Resolve-RepositoryPath -RepositoryRoot $SolutionDir -PathValue $option.Value
    $allowedRoot = Join-Path $SolutionDir "artifacts/$($option.Root)"
    if ([string]::IsNullOrWhiteSpace($resolved) -or -not (Test-PathWithinRoot -CandidatePath $resolved -RootPath $allowedRoot)) {
        throw "$($option.Name) must be within $allowedRoot. External/user-owned directories are not cleanup targets."
    }
    $targets.Add($resolved)
}

$trackedPaths = @(Get-TrackedRepositoryPaths -RepositoryRoot $SolutionDir)
# Validate every target before deleting any output.
$validatedTargets = @($targets | Sort-Object -Unique | ForEach-Object {
    Assert-RepositoryCleanupPath -RepositoryRoot $SolutionDir -PathValue $_ -TrackedPaths $trackedPaths
})

Write-Step 'Cleaning generated repository outputs'
$deletedCount = 0
foreach ($target in $validatedTargets) {
    if (-not (Test-Path -LiteralPath $target)) { continue }
    if ($DryRun) { Write-Info "[DRY-RUN] Would delete: $target" }
    else {
        Remove-Item -LiteralPath $target -Recurse -Force -ErrorAction Stop
        if (Test-Path -LiteralPath $target) { throw "Cleanup incomplete: $target" }
        Write-Info "Removed: $target"
    }
    $deletedCount++
}

if ($All) {
    if ($DryRun) { Write-Info '[DRY-RUN] Would clear NuGet HTTP cache' }
    else { Invoke-DotNetCommand -Arguments @('nuget','locals','http-cache','--clear') -ErrorMessage 'NuGet HTTP cache cleanup failed' }
}
if (-not $SkipDotNetClean) {
    $solution = Get-SolutionPath -RepositoryRoot $SolutionDir
    if (-not (Test-Path -LiteralPath $solution -PathType Leaf)) { throw "Solution not found: $solution" }
    if ($DryRun) { Write-Info "[DRY-RUN] Would run dotnet clean $solution" }
    else { Invoke-DotNetCommand -Arguments @('clean',$solution,'--verbosity','minimal') -ErrorMessage 'dotnet clean failed' }
}
Write-Success "Clean completed: $deletedCount generated directory target(s). Reports, personal state and default export/import directories preserved."
