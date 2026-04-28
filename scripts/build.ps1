#Requires -Version 7.0

[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",

    [switch]$Clean = $true,

    [switch]$LockedMode = $false,

    [ValidateNotNullOrEmpty()]
    [string]$RuntimeIdentifier = "win-x64",

    [switch]$SelfContained = $false
)

. (Join-Path $PSScriptRoot "internal/common.ps1")

$repositoryRoot = Get-RepositoryRoot -ScriptRoot $PSScriptRoot
$scriptPath = Join-Path $repositoryRoot "build\\build.ps1"

Write-Step "Running canonical build"
Invoke-RepositoryPowerShellScript `
    -ScriptPath $scriptPath `
    -Arguments @(
        "-Configuration", $Configuration,
        "-Clean:$([bool]$Clean)",
        "-Publish:$false",
        "-LockedMode:$([bool]$LockedMode)",
        "-RuntimeIdentifier", $RuntimeIdentifier,
        "-SelfContained:$([bool]$SelfContained)"
    ) `
    -ErrorMessage "canonical build failed"

Write-Success "Build completed"
