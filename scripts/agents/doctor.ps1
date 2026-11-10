#Requires -Version 7.0

[CmdletBinding()]
param(
    [switch]$CheckPackaging
)

. (Join-Path $PSScriptRoot "../internal/common.ps1")

$repositoryRoot = Get-RepositoryRoot -ScriptRoot $PSScriptRoot
$solutionPath = Get-SolutionPath -RepositoryRoot $repositoryRoot
$globalJsonPath = Join-Path $repositoryRoot "global.json"

Write-Step "Validating local toolchain"
Assert-WindowsPlatform
$sdkSpecification = Assert-DotNetSdkPinnedVersion -RepositoryRoot $repositoryRoot
Assert-CommandAvailable -CommandName "pwsh"

if (-not (Test-Path $solutionPath -PathType Leaf)) {
    throw "Solution file not found: $solutionPath"
}

if (-not (Test-Path $globalJsonPath -PathType Leaf)) {
    throw "global.json not found: $globalJsonPath"
}

Write-Info ".NET SDK: $($sdkSpecification.Version)"
Write-Info "PowerShell: $($PSVersionTable.PSVersion)"
Write-Info "App version: $(Get-ApplicationVersion -RepositoryRoot $repositoryRoot)"
Write-Info "Solution: $solutionPath"

if ($CheckPackaging) {
    $innoSetupBinPath = Get-InnoSetupBinPath -RepositoryRoot $repositoryRoot
    if ($null -eq $innoSetupBinPath) {
        throw "Inno Setup 6 not found. Install it or set INNOSETUP_BIN."
    }

    $isccPath = Get-InnoSetupCompilerPath -RepositoryRoot $repositoryRoot
    Write-Info "Inno Setup bin: $innoSetupBinPath"
    Write-Info "ISCC: $isccPath"
}

Write-Success "Doctor checks passed"
