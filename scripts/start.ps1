#Requires -Version 7.0

[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug",

    [ValidateNotNullOrEmpty()]
    [string]$RuntimeIdentifier = "win-x64",

    [switch]$NoBuild
)

. (Join-Path $PSScriptRoot "internal/common.ps1")

$repositoryRoot = Get-RepositoryRoot -ScriptRoot $PSScriptRoot
$shellProjectPath = Join-Path $repositoryRoot "src/OnlyExo365.Shell/OnlyExo365.Shell.csproj"

if (-not (Test-Path $shellProjectPath -PathType Leaf)) {
    throw "Shell project not found: $shellProjectPath"
}

Write-Step "Starting OnlyExo365"

$arguments = @(
    "run",
    "--project",
    $shellProjectPath,
    "--configuration",
    $Configuration,
    "--runtime",
    $RuntimeIdentifier
)

if ($NoBuild) {
    $arguments += "--no-build"
}

Invoke-DotNetCommand -Arguments $arguments -ErrorMessage "OnlyExo365 startup failed"
