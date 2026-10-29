#Requires -Version 7.0

[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$CompileConfiguration = "Debug",

    [ValidateSet("Release")]
    [string]$PackageConfiguration = "Release",

    [bool]$LockedMode = $true,

    [Alias("RuntimeIdentifier")]
    [ValidateNotNullOrEmpty()]
    [string[]]$RuntimeIdentifiers = @("win-x64"),

    [switch]$SelfContained,

    [switch]$InstallPrerequisites,

    [ValidateSet("Auto", "Winget", "Chocolatey")]
    [string]$PackageManager = "Auto",

    [switch]$SkipTests,

    [switch]$SkipSecurityScans,

    [switch]$SkipPackaging,

    [switch]$RunReproducibility,

    [switch]$RunSigningValidation
)

. (Join-Path $PSScriptRoot "internal/common.ps1")

$repositoryRoot = Get-RepositoryRoot -ScriptRoot $PSScriptRoot
$solutionPath = Get-SolutionPath -RepositoryRoot $repositoryRoot
$resolvedRuntimeIdentifiers = Resolve-RuntimeIdentifiers -RequestedRuntimeIdentifiers $RuntimeIdentifiers -DefaultRuntimeIdentifiers @("win-x64")
$primaryRuntimeIdentifier = @($resolvedRuntimeIdentifiers)[0]
$startedAt = Get-Date

function Invoke-GateScript {
    param(
        [string]$RelativePath,
        [string[]]$Arguments,
        [string]$ErrorMessage
    )

    Invoke-RepositoryPowerShellScript `
        -ScriptPath (Join-Path $repositoryRoot $RelativePath) `
        -Arguments $Arguments `
        -ErrorMessage $ErrorMessage
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " OnlyExo365 Local Gate" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Repository      : $repositoryRoot"
Write-Host "Compile config  : $CompileConfiguration"
Write-Host "Package config  : $PackageConfiguration"
Write-Host "Runtime IDs     : $($resolvedRuntimeIdentifiers -join ', ')"
Write-Host "Locked restore  : $LockedMode"
Write-Host "Install prereqs : $($InstallPrerequisites.IsPresent)"
Write-Host "Started at      : $($startedAt.ToString('HH:mm:ss'))"

Write-Step "Cleaning repository outputs"
Invoke-GateScript `
    -RelativePath "scripts\clean.ps1" `
    -Arguments @("-All") `
    -ErrorMessage "repository clean failed"

Write-Step "Checking prerequisites"
Assert-WindowsPlatform
Assert-CommandAvailable -CommandName "pwsh"
Assert-DotNetSdkPinnedVersion -RepositoryRoot $repositoryRoot | Out-Null

$innoArguments = @()
if ($InstallPrerequisites) {
    $innoArguments += "-Install", "-PackageManager", $PackageManager
}

Invoke-GateScript `
    -RelativePath "scripts\Install-InnoSetup.ps1" `
    -Arguments $innoArguments `
    -ErrorMessage "Inno Setup prerequisite check failed"

Invoke-GateScript `
    -RelativePath "scripts\agents\doctor.ps1" `
    -Arguments @("-CheckPackaging") `
    -ErrorMessage "doctor prerequisite check failed"

Write-Step "Bootstrapping dependencies"
Invoke-GateScript `
    -RelativePath "scripts\bootstrap.ps1" `
    -Arguments @("-LockedMode:$LockedMode", "-RuntimeIdentifiers", ($resolvedRuntimeIdentifiers -join ',')) `
    -ErrorMessage "bootstrap failed"

Write-Step "Compiling solution"
Invoke-GateScript `
    -RelativePath "scripts\agents\compile.ps1" `
    -Arguments @("-Configuration", $CompileConfiguration, "-LockedMode:$LockedMode", "-RuntimeIdentifiers", ($resolvedRuntimeIdentifiers -join ','), "-NoBootstrap") `
    -ErrorMessage "compile gate failed"

if ($SkipTests) {
    Write-Step "Running tests"
    Write-Info "Tests skipped by -SkipTests"
}
else {
    Write-Step "Running tests"
    Invoke-GateScript `
        -RelativePath "scripts\agents\test.ps1" `
        -Arguments @("-Configuration", $CompileConfiguration, "-LockedMode:$LockedMode", "-RuntimeIdentifier", $primaryRuntimeIdentifier, "-ResultsDirectory", "artifacts/test-results/unit", "-NoBootstrap", "-NoBuild", "-NoRestore") `
        -ErrorMessage "test gate failed"
}

if ($SkipSecurityScans) {
    Write-Step "Running repository validation"
    Write-Info "Security and architecture scans skipped by -SkipSecurityScans"
}
else {
    Write-Step "Running repository validation"
    New-Item -Path (Join-Path $repositoryRoot "artifacts\security") -ItemType Directory -Force | Out-Null

    Invoke-GateScript `
        -RelativePath "build\assert-architecture-constraints.ps1" `
        -Arguments @() `
        -ErrorMessage "architecture gate failed"

    Invoke-GateScript `
        -RelativePath "build\assert-no-vulnerable-packages.ps1" `
        -Arguments @("-SolutionPath", $solutionPath, "-ReportPath", "artifacts/security/nuget-vulnerabilities.json") `
        -ErrorMessage "NuGet vulnerability gate failed"

    Invoke-GateScript `
        -RelativePath "build\run-secret-scan.ps1" `
        -Arguments @("-SourcePath", ".", "-ReportPath", "artifacts/security/gitleaks.sarif") `
        -ErrorMessage "secret scan gate failed"
}

if ($SkipPackaging) {
    Write-Step "Building distributable artifacts"
    Write-Info "Packaging skipped by -SkipPackaging"
}
else {
    Write-Step "Building distributable artifacts"

    $packArguments = @(
        "-Configuration", $PackageConfiguration,
        "-LockedMode:$LockedMode",
        "-Clean:$false",
        "-RuntimeIdentifiers", ($resolvedRuntimeIdentifiers -join ','),
        "-SelfContained:$([bool]$SelfContained)"
    )

    Invoke-GateScript `
        -RelativePath "scripts\pack.ps1" `
        -Arguments $packArguments `
        -ErrorMessage "packaging gate failed"

    if ($RunReproducibility) {
        Invoke-GateScript `
            -RelativePath "build\verify-package-reproducibility.ps1" `
            -Arguments @("-Configuration", $PackageConfiguration, "-RuntimeIdentifier", $primaryRuntimeIdentifier) `
            -ErrorMessage "package reproducibility gate failed"
    }
    else {
        Write-Info "Package reproducibility skipped. Use -RunReproducibility to enable it."
    }

    if ($RunSigningValidation) {
        Invoke-GateScript `
            -RelativePath "build\validate-artifact-signing.ps1" `
            -Arguments @("-Path", "artifacts/publish", "artifacts/packages") `
            -ErrorMessage "artifact signing validation gate failed"
    }
    else {
        Write-Info "Signing validation skipped. Use -RunSigningValidation to enable it."
    }
}

$endedAt = Get-Date
$duration = $endedAt - $startedAt

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " LOCAL GATE COMPLETED" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Duration: $($duration.TotalSeconds.ToString('F1')) seconds" -ForegroundColor Gray

if (-not $SkipPackaging) {
    Write-Host "Publish : $(Get-PublishArtifactsPath -RepositoryRoot $repositoryRoot)" -ForegroundColor Cyan
    Write-Host "Package : $(Join-Path $repositoryRoot 'artifacts\packages\OnlyExo365.Setup.exe')" -ForegroundColor Cyan
}
