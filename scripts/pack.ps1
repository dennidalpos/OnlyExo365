#Requires -Version 7.0

[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",

    [switch]$LockedMode = $true,

    [switch]$Clean = $false,

    [Alias("RuntimeIdentifier")]
    [ValidateNotNullOrEmpty()]
    [string[]]$RuntimeIdentifiers = @("win-x64", "win-x86"),

    [switch]$SelfContained = $false
)

. (Join-Path $PSScriptRoot "helpers/common.ps1")

$repositoryRoot = Get-RepositoryRoot -ScriptRoot $PSScriptRoot
$buildScriptPath = Join-Path $repositoryRoot "build\\build.ps1"
$setupExeScriptPath = Join-Path $repositoryRoot "build\\create-setup-exe.ps1"
$packagesDirectory = Join-Path $repositoryRoot "artifacts\\packages"
$canonicalSetupExePath = Join-Path $packagesDirectory "OnlyExo365.Setup.exe"
$publishDirectory = Get-PublishArtifactsPath -RepositoryRoot $repositoryRoot
$resolvedRuntimeIdentifiers = Resolve-RuntimeIdentifiers -RequestedRuntimeIdentifiers $RuntimeIdentifiers -DefaultRuntimeIdentifiers @("win-x64", "win-x86")

foreach ($requiredRuntimeIdentifier in @("win-x64", "win-x86")) {
    if ($resolvedRuntimeIdentifiers -notcontains $requiredRuntimeIdentifier) {
        $resolvedRuntimeIdentifiers += $requiredRuntimeIdentifier
    }
}

Write-Step "Packaging distributable artifacts"
Initialize-ArtifactsLayout -RepositoryRoot $repositoryRoot | Out-Null

$buildArguments = @(
    "-Configuration", $Configuration,
    "-Clean:$([bool]$Clean)",
    "-Publish:$true",
    "-LockedMode:$([bool]$LockedMode)",
    "-RuntimeIdentifiers", ($resolvedRuntimeIdentifiers -join ',')
) + @(
    "-SelfContained:$([bool]$SelfContained)"
)

Invoke-RepositoryPowerShellScript `
    -ScriptPath $buildScriptPath `
    -Arguments $buildArguments `
    -ErrorMessage "pack failed"

if (-not (Test-Path $publishDirectory -PathType Container)) {
    throw "Expected publish output not found: $publishDirectory"
}

Invoke-RepositoryPowerShellScript `
    -ScriptPath $setupExeScriptPath `
    -Arguments @(
        "-PublishPath", $publishDirectory,
        "-OutputDirectory", $packagesDirectory,
        "-OutputFileName", "OnlyExo365.Setup.exe",
        "-ProductName", "OnlyExo365"
    ) `
    -ErrorMessage "setup exe packaging failed"

if (-not (Test-Path $canonicalSetupExePath -PathType Leaf)) {
    throw "Expected setup EXE not found: $canonicalSetupExePath"
}

Write-Info "Canonical setup EXE: $canonicalSetupExePath"
Write-Success "Packaging completed"
