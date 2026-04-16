#Requires -Version 7.0

[CmdletBinding()]
param(
    [ValidateSet("Release")]
    [string]$Configuration = "Release",

    [switch]$SelfContained = $true,

    [ValidateNotNullOrEmpty()]
    [string]$RuntimeIdentifier = "win-x64",

    [switch]$RunTests = $true,

    [switch]$CreateTag = $false,

    [string]$TagName,

    [string]$ReleaseName,

    [string]$OutputRoot = "artifacts/release"
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptDir
$solutionFile = Join-Path $repoRoot "ExchangeAdmin.sln"
$buildScript = Join-Path $scriptDir "build.ps1"
$presentationProject = Join-Path $repoRoot "src\ExchangeAdmin.Presentation\ExchangeAdmin.Presentation.csproj"
$outputRootPath = if ([System.IO.Path]::IsPathRooted($OutputRoot)) { $OutputRoot } else { Join-Path $repoRoot $OutputRoot }

. (Join-Path $scriptDir "helpers\common.ps1")

function Invoke-Native {
    param(
        [string]$FilePath,
        [string[]]$Arguments,
        [string]$ErrorMessage
    )

    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) {
        Stop-WithError "$ErrorMessage (exit code: $LASTEXITCODE)" $LASTEXITCODE
    }
}

function Invoke-DotNet {
    param(
        [string]$Command,
        [string[]]$Arguments,
        [string]$ErrorMessage
    )

    Invoke-Native -FilePath "dotnet" -Arguments (@($Command) + $Arguments) -ErrorMessage $ErrorMessage
}

function Get-RequiredCommand {
    param([string]$Name)

    $command = Get-Command $Name -ErrorAction SilentlyContinue
    if ($null -eq $command) {
        Stop-WithError "Required command not found in PATH: $Name"
    }

    return $command.Source
}

function Get-GitStatusLines {
    $statusOutput = & git -C $repoRoot status --porcelain=v1 --untracked-files=all
    if ($LASTEXITCODE -ne 0) {
        Stop-WithError "Unable to read git status."
    }

    return @($statusOutput | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

function New-FileHashRecord {
    param(
        [string]$BasePath,
        [System.IO.FileInfo]$File
    )

    $hash = Get-FileHash -Algorithm SHA256 -Path $File.FullName
    return [ordered]@{
        path = [System.IO.Path]::GetRelativePath($BasePath, $File.FullName).Replace('\', '/')
        size_bytes = $File.Length
        sha256 = $hash.Hash.ToLowerInvariant()
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " OnlyExo365 Baseline Release" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

Get-RequiredCommand -Name "git" | Out-Null
Get-RequiredCommand -Name "dotnet" | Out-Null
Get-RequiredCommand -Name "pwsh" | Out-Null

if (-not (Test-Path $solutionFile)) {
    Stop-WithError "Solution file not found: $solutionFile"
}

if (-not (Test-Path $buildScript)) {
    Stop-WithError "Build script not found: $buildScript"
}

if (-not (Test-Path $presentationProject)) {
    Stop-WithError "Presentation project not found: $presentationProject"
}

Write-Step "Validating repository state"

$statusLines = Get-GitStatusLines
if ($statusLines.Count -gt 0) {
    Write-Host "   Dirty worktree detected:" -ForegroundColor Yellow
    $statusLines | ForEach-Object { Write-Host "   $_" -ForegroundColor Yellow }
    Stop-WithError "Baseline release requires a clean git worktree."
}

$commit = (& git -C $repoRoot rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($commit)) {
    Stop-WithError "Unable to resolve current commit."
}

$branch = (& git -C $repoRoot branch --show-current).Trim()
if ($LASTEXITCODE -ne 0) {
    Stop-WithError "Unable to resolve current branch."
}

$version = (& dotnet msbuild $presentationProject "-nologo" "-getProperty:Version").Trim()
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($version)) {
    Stop-WithError "Unable to resolve application version from MSBuild."
}

$resolvedReleaseName = if ([string]::IsNullOrWhiteSpace($ReleaseName)) {
    "OnlyExo365-$version-$RuntimeIdentifier"
}
else {
    $ReleaseName.Trim()
}

$resolvedTagName = if ([string]::IsNullOrWhiteSpace($TagName)) {
    "baseline/v$version"
}
else {
    $TagName.Trim()
}

$releaseDir = Join-Path $outputRootPath $resolvedReleaseName
$packagePath = Join-Path $releaseDir "$resolvedReleaseName.zip"
$manifestPath = Join-Path $releaseDir "baseline-manifest.json"
$summaryPath = Join-Path $releaseDir "baseline-summary.txt"
$publishDir = Join-Path $repoRoot "artifacts\publish"

Write-Info "Commit     : $commit"
Write-Info "Branch     : $branch"
Write-Info "Version    : $version"
Write-Info "Release    : $resolvedReleaseName"
Write-Info "Tag        : $resolvedTagName"
Write-Info "OutputRoot : $outputRootPath"

if (Test-Path $releaseDir) {
    Remove-Item -Path $releaseDir -Recurse -Force
}

New-Item -Path $releaseDir -ItemType Directory -Force | Out-Null

if ($RunTests) {
    Write-Step "Running automated tests"
    Invoke-DotNet -Command "test" -Arguments @($solutionFile, "-c", "Debug", "--nologo", "--verbosity", "minimal") -ErrorMessage "Test run failed"
    Write-Success "Test suite passed"
}
else {
    Write-Info "Test execution skipped by parameter"
}

Write-Step "Building publish output"

$buildArguments = @(
    "-File", $buildScript,
    "-Configuration", $Configuration,
    "-Clean",
    "-Publish",
    "-RuntimeIdentifier", $RuntimeIdentifier
)

if ($SelfContained) {
    $buildArguments += "-SelfContained"
}
else {
    $buildArguments += "-SelfContained:`$false"
}

Invoke-Native -FilePath "pwsh" -Arguments $buildArguments -ErrorMessage "Build script failed"
Write-Success "Build script completed"

if (-not (Test-Path $publishDir)) {
    Stop-WithError "Publish output not found: $publishDir"
}

if (-not (Test-Path $releaseDir)) {
    New-Item -Path $releaseDir -ItemType Directory -Force | Out-Null
}

Write-Step "Collecting release evidence"

$publishFiles = Get-ChildItem -Path $publishDir -Recurse -File | Sort-Object FullName
if ($publishFiles.Count -eq 0) {
    Stop-WithError "Publish directory is empty: $publishDir"
}

$fileEntries = @($publishFiles | ForEach-Object { New-FileHashRecord -BasePath $publishDir -File $_ })
Compress-Archive -Path (Join-Path $publishDir "*") -DestinationPath $packagePath -CompressionLevel Optimal
$packageHash = (Get-FileHash -Algorithm SHA256 -Path $packagePath).Hash.ToLowerInvariant()

$manifest = [ordered]@{
    generated_at_utc = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
    release_name = $resolvedReleaseName
    version = $version
    configuration = $Configuration
    runtime_identifier = $RuntimeIdentifier
    self_contained = [bool]$SelfContained
    git = [ordered]@{
        branch = $branch
        commit = $commit
        tag = if ($CreateTag) { $resolvedTagName } else { $null }
        worktree_clean = $true
    }
    verification = [ordered]@{
        tests_ran = [bool]$RunTests
        publish_directory = [System.IO.Path]::GetRelativePath($repoRoot, $publishDir).Replace('\', '/')
        published_files = $fileEntries
        zip = [ordered]@{
            path = [System.IO.Path]::GetRelativePath($repoRoot, $packagePath).Replace('\', '/')
            size_bytes = (Get-Item $packagePath).Length
            sha256 = $packageHash
        }
    }
}

$manifest | ConvertTo-Json -Depth 6 | Set-Content -Path $manifestPath -Encoding UTF8

$summaryLines = @(
    "OnlyExo365 baseline release summary",
    "GeneratedAtUtc : $($manifest.generated_at_utc)",
    "ReleaseName    : $resolvedReleaseName",
    "Version        : $version",
    "Commit         : $commit",
    "Branch         : $branch",
    "Runtime        : $RuntimeIdentifier",
    "SelfContained  : $SelfContained",
    "TestsRan       : $RunTests",
    "ZipPath        : $([System.IO.Path]::GetRelativePath($repoRoot, $packagePath).Replace('\', '/'))",
    "ZipSha256      : $packageHash",
    "PublishedFiles : $($fileEntries.Count)"
)
$summaryLines | Set-Content -Path $summaryPath -Encoding UTF8

Write-Success "Manifest created: $manifestPath"
Write-Success "Summary created : $summaryPath"
Write-Success "Package created : $packagePath"

if ($CreateTag) {
    Write-Step "Creating git tag"

    $existingTagOutput = & git -C $repoRoot tag --list $resolvedTagName
    if ($LASTEXITCODE -ne 0) {
        Stop-WithError "Unable to inspect existing tags."
    }

    $existingTag = if ($null -eq $existingTagOutput) { "" } else { ($existingTagOutput | Out-String).Trim() }

    if (-not [string]::IsNullOrWhiteSpace($existingTag)) {
        Stop-WithError "Tag already exists: $resolvedTagName"
    }

    Invoke-Native -FilePath "git" -Arguments @("-C", $repoRoot, "tag", "-a", $resolvedTagName, "-m", "Baseline release $resolvedReleaseName") -ErrorMessage "Tag creation failed"
    Write-Success "Tag created: $resolvedTagName"
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " BASELINE READY" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Manifest: $manifestPath" -ForegroundColor Cyan
Write-Host "Package : $packagePath" -ForegroundColor Cyan
Write-Host ""
