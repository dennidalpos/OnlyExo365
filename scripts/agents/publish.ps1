#Requires -Version 7.0

[CmdletBinding()]
param(
    [string]$ReleaseTag,

    [string]$PublishPath = "artifacts/publish",

    [string]$SetupExePath = "artifacts/packages/OnlyExo365.Setup.exe",

    [ValidateNotNullOrEmpty()]
    [string]$RuntimeIdentifier = "win-x64",

    [string]$OutputDirectory = "artifacts/publish/release-assets"
)

. (Join-Path $PSScriptRoot "../internal/common.ps1")

$repositoryRoot = Get-RepositoryRoot -ScriptRoot $PSScriptRoot
$publishScriptPath = Join-Path $repositoryRoot "build\\publish-release-assets.ps1"
$resolvedPublishPath = Resolve-RepositoryPath -RepositoryRoot $repositoryRoot -PathValue $PublishPath
$resolvedSetupExePath = Resolve-RepositoryPath -RepositoryRoot $repositoryRoot -PathValue $SetupExePath
$applicationVersion = Get-ApplicationVersion -RepositoryRoot $repositoryRoot

if ([string]::IsNullOrWhiteSpace($ReleaseTag)) {
    throw "ReleaseTag is required. Current application version: $applicationVersion. Pass -ReleaseTag v$applicationVersion (or the next release tag) explicitly."
}

if (-not (Test-Path $resolvedPublishPath -PathType Container)) {
    throw "Publish payload not found at $resolvedPublishPath. Run scripts/pack.ps1 first."
}

if (-not (Test-Path $resolvedSetupExePath -PathType Leaf)) {
    throw "Setup EXE not found at $SetupExePath. Run scripts/pack.ps1 first."
}

Write-Step "Publishing packaged artifacts"
Invoke-RepositoryPowerShellScript `
    -ScriptPath $publishScriptPath `
    -Arguments @(
        "-PublishPath", $resolvedPublishPath,
        "-SetupExePath", $resolvedSetupExePath,
        "-ReleaseTag", $ReleaseTag,
        "-RuntimeIdentifier", $RuntimeIdentifier,
        "-OutputDirectory", (Resolve-RepositoryPath -RepositoryRoot $repositoryRoot -PathValue $OutputDirectory)
    ) `
    -ErrorMessage "publish failed"

Write-Success "Publish assets prepared"
