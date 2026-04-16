#Requires -Version 7.0

[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug",

    [bool]$LockedMode = $true,

    [Alias("RuntimeIdentifier")]
    [ValidateNotNullOrEmpty()]
    [string[]]$RuntimeIdentifiers = @("win-x64", "win-x86"),

    [switch]$NoBootstrap
)

. (Join-Path $PSScriptRoot "helpers/common.ps1")

$repositoryRoot = Get-RepositoryRoot -ScriptRoot $PSScriptRoot
$solutionPath = Get-SolutionPath -RepositoryRoot $repositoryRoot
$buildArtifactsPath = Get-BuildArtifactsPath -RepositoryRoot $repositoryRoot
$resolvedRuntimeIdentifiers = Resolve-RuntimeIdentifiers -RequestedRuntimeIdentifiers $RuntimeIdentifiers -DefaultRuntimeIdentifiers @("win-x64", "win-x86")

if (-not $NoBootstrap) {
    $bootstrapArguments = @("-LockedMode:$LockedMode", "-RuntimeIdentifiers", ($resolvedRuntimeIdentifiers -join ','))
    Invoke-RepositoryPowerShellScript `
        -ScriptPath (Join-Path $PSScriptRoot "bootstrap.ps1") `
        -Arguments $bootstrapArguments `
        -ErrorMessage "bootstrap failed"
}

Write-Step "Compiling solution"
Initialize-ArtifactsLayout -RepositoryRoot $repositoryRoot | Out-Null

$arguments = @(
    "build",
    $solutionPath,
    "-c",
    $Configuration,
    "--no-restore",
    "--artifacts-path",
    $buildArtifactsPath,
    "/warnaserror"
)

Invoke-DotNetCommand -Arguments $arguments -ErrorMessage "dotnet build failed"
Write-Success "Compilation completed"
