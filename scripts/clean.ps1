#Requires -Version 7.0

[CmdletBinding()]
param(
    [switch]$All,
    [switch]$DryRun,
    [switch]$SkipDotNetClean,
    [string]$ExportDirPath,
    [string]$ImportDirPath
)

. (Join-Path $PSScriptRoot "helpers/common.ps1")

$repositoryRoot = Get-RepositoryRoot -ScriptRoot $PSScriptRoot
$scriptPath = Join-Path $repositoryRoot "build\\clean.ps1"
$arguments = @()

if ($All) {
    $arguments += "-All"
}

if ($DryRun) {
    $arguments += "-DryRun"
}

if ($SkipDotNetClean) {
    $arguments += "-SkipDotNetClean"
}

if (-not [string]::IsNullOrWhiteSpace($ExportDirPath)) {
    $arguments += "-ExportDirPath", $ExportDirPath
}

if (-not [string]::IsNullOrWhiteSpace($ImportDirPath)) {
    $arguments += "-ImportDirPath", $ImportDirPath
}

Write-Step "Running canonical clean"
Invoke-RepositoryPowerShellScript -ScriptPath $scriptPath -Arguments $arguments -ErrorMessage "clean failed"
Write-Success "Clean completed"
