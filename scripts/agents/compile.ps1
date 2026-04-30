#Requires -Version 7.0

[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug",

    [bool]$LockedMode = $true,

    [Alias("RuntimeIdentifier")]
    [ValidateNotNullOrEmpty()]
    [string[]]$RuntimeIdentifiers = @("win-x64"),

    [switch]$NoBootstrap
)

. (Join-Path $PSScriptRoot "../internal/common.ps1")

$repositoryRoot = Get-RepositoryRoot -ScriptRoot $PSScriptRoot
$resolvedRuntimeIdentifiers = Resolve-RuntimeIdentifiers -RequestedRuntimeIdentifiers $RuntimeIdentifiers -DefaultRuntimeIdentifiers @("win-x64")

Invoke-RepositoryBootstrap `
    -RepositoryRoot $repositoryRoot `
    -LockedMode $LockedMode `
    -RuntimeIdentifiers $resolvedRuntimeIdentifiers `
    -Skip:$NoBootstrap

Write-Step "Compiling solution"
Invoke-RepositoryPowerShellScript `
    -ScriptPath (Join-Path $repositoryRoot "scripts/build.ps1") `
    -Arguments @(
        "-Configuration", $Configuration,
        "-Clean:$false",
        "-LockedMode:$false",
        "-RuntimeIdentifier", (@($resolvedRuntimeIdentifiers)[0]),
        "-NoRestore:$true"
    ) `
    -ErrorMessage "canonical build failed"
Write-Success "Compilation completed"
