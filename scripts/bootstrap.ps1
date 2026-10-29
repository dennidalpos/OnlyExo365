#Requires -Version 7.0

[CmdletBinding()]
param(
    [bool]$LockedMode = $true,

    [Alias("RuntimeIdentifier")]
    [ValidateNotNullOrEmpty()]
    [string[]]$RuntimeIdentifiers = @("win-x64")
)

. (Join-Path $PSScriptRoot "internal/common.ps1")

$repositoryRoot = Get-RepositoryRoot -ScriptRoot $PSScriptRoot
$solutionPath = Get-SolutionPath -RepositoryRoot $repositoryRoot
$buildArtifactsPath = Get-BuildArtifactsPath -RepositoryRoot $repositoryRoot
$resolvedRuntimeIdentifiers = Resolve-RuntimeIdentifiers -RequestedRuntimeIdentifiers $RuntimeIdentifiers -DefaultRuntimeIdentifiers @("win-x64")

Write-Step "Bootstrapping repository prerequisites"
Assert-WindowsPlatform
$sdkSpecification = Assert-DotNetSdkPinnedVersion -RepositoryRoot $repositoryRoot
Initialize-ArtifactsLayout -RepositoryRoot $repositoryRoot | Out-Null

if (-not (Test-Path $solutionPath -PathType Leaf)) {
    throw "Solution file not found: $solutionPath"
}

Write-Info ".NET SDK: $($sdkSpecification.Version)"
Write-Info "Runtime identifiers: $($resolvedRuntimeIdentifiers -join ', ')"
Write-Info "Locked restore: $LockedMode"

$arguments = @(
    "restore",
    $solutionPath,
    "--verbosity",
    "minimal",
    "--artifacts-path",
    $buildArtifactsPath
)

if ($LockedMode) {
    $arguments += "--locked-mode"
}

Write-Info "Restoring with project runtime graph: $($resolvedRuntimeIdentifiers -join ', ')"
Invoke-DotNetCommand -Arguments $arguments -ErrorMessage "dotnet restore failed"

Write-Success "Repository bootstrap completed"
