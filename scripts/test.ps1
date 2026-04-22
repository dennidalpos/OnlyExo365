#Requires -Version 7.0

[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug",

    [bool]$LockedMode = $true,

    [ValidateNotNullOrEmpty()]
    [string]$RuntimeIdentifier = "win-x64",

    [string]$ResultsDirectory = "artifacts/test-results",

    [switch]$NoBootstrap,

    [switch]$NoBuild,

    [switch]$NoRestore
)

. (Join-Path $PSScriptRoot "helpers/common.ps1")

$repositoryRoot = Get-RepositoryRoot -ScriptRoot $PSScriptRoot
$solutionPath = Get-SolutionPath -RepositoryRoot $repositoryRoot
$resolvedResultsDirectory = Resolve-RepositoryPath -RepositoryRoot $repositoryRoot -PathValue $ResultsDirectory
$buildArtifactsPath = Get-BuildArtifactsPath -RepositoryRoot $repositoryRoot

if (-not $NoBootstrap) {
    Invoke-RepositoryPowerShellScript `
        -ScriptPath (Join-Path $PSScriptRoot "bootstrap.ps1") `
        -Arguments @("-LockedMode:$LockedMode", "-RuntimeIdentifier", $RuntimeIdentifier) `
        -ErrorMessage "bootstrap failed"
}

Write-Step "Running automated tests"
New-Item -Path $resolvedResultsDirectory -ItemType Directory -Force | Out-Null

$arguments = @(
    "test",
    $solutionPath,
    "-c",
    $Configuration,
    "--logger",
    "trx;LogFileName=onlyexo365-tests.trx",
    "--results-directory",
    $resolvedResultsDirectory,
    "--artifacts-path",
    $buildArtifactsPath,
    "--disable-build-servers",
    "--verbosity",
    "minimal"
)

if ($NoBuild) {
    $arguments += "--no-build"
}

if ($NoRestore -or (-not $NoBootstrap)) {
    $arguments += "--no-restore"
}

Invoke-DotNetCommand -Arguments $arguments -ErrorMessage "dotnet test failed"
Write-Success "Tests completed"
