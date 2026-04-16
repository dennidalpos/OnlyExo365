#Requires -Version 7.0

[CmdletBinding()]
param(
    [string]$SolutionPath = "ExchangeAdmin.sln",
    [switch]$IncludeTransitive = $true,
    [string]$ReportPath = "artifacts/security/nuget-vulnerabilities.json",
    [switch]$LockedMode = $true
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$repositoryRoot = Split-Path -Parent $scriptDirectory

. (Join-Path $scriptDirectory "helpers\common.ps1")

function Resolve-RepoPath {
    param(
        [string]$BaseDirectory,
        [string]$PathValue
    )

    if ([System.IO.Path]::IsPathRooted($PathValue)) {
        return $PathValue
    }

    return Join-Path $BaseDirectory $PathValue
}

function Get-ObjectPropertyValue {
    param(
        [object]$Object,
        [string]$PropertyName
    )

    if ($null -eq $Object) {
        return $null
    }

    $property = $Object.PSObject.Properties[$PropertyName]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function Get-SolutionProjectPaths {
    param([string]$SolutionPath)

    $rawProjects = & dotnet sln $SolutionPath list
    if ($LASTEXITCODE -ne 0) {
        Stop-WithError "dotnet sln list failed for $SolutionPath (exit code: $LASTEXITCODE)" $LASTEXITCODE
    }

    $projectPaths = @(
        $rawProjects |
            Where-Object { $_ -match '\.csproj$' } |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { Resolve-RepoPath -BaseDirectory $repositoryRoot -PathValue $_ }
    )

    if ($projectPaths.Count -eq 0) {
        Stop-WithError "No C# project was found in solution $SolutionPath."
    }

    return $projectPaths
}

function Invoke-SolutionRestoreForPackageInspection {
    param(
        [string]$SolutionPath,
        [bool]$UseLockedMode
    )

    $restoreArguments = @(
        "restore",
        $SolutionPath,
        "--verbosity",
        "minimal"
    )

    if ($UseLockedMode) {
        $restoreArguments += "--locked-mode"
    }

    & dotnet @restoreArguments
    if ($LASTEXITCODE -ne 0) {
        Stop-WithError "dotnet restore failed before vulnerability inspection (exit code: $LASTEXITCODE)" $LASTEXITCODE
    }
}

function Get-VulnerabilityRecords {
    param([object]$Payload)

    $records = @()

    foreach ($project in @(Get-ObjectPropertyValue -Object $Payload -PropertyName "projects")) {
        if ($null -eq $project) {
            continue
        }

        $projectPath = [string](Get-ObjectPropertyValue -Object $project -PropertyName "path")

        foreach ($framework in @(Get-ObjectPropertyValue -Object $project -PropertyName "frameworks")) {
            if ($null -eq $framework) {
                continue
            }

            $frameworkName = [string](Get-ObjectPropertyValue -Object $framework -PropertyName "framework")

            foreach ($packageBucket in @(
                @{ Name = "topLevel"; Packages = (Get-ObjectPropertyValue -Object $framework -PropertyName "topLevelPackages") },
                @{ Name = "transitive"; Packages = (Get-ObjectPropertyValue -Object $framework -PropertyName "transitivePackages") }
            )) {
                foreach ($package in @($packageBucket.Packages)) {
                    if ($null -eq $package) {
                        continue
                    }

                    foreach ($vulnerability in @(Get-ObjectPropertyValue -Object $package -PropertyName "vulnerabilities")) {
                        if ($null -eq $vulnerability) {
                            continue
                        }

                        $records += [pscustomobject]@{
                            Project = $projectPath
                            Framework = $frameworkName
                            PackageType = $packageBucket.Name
                            PackageId = [string](Get-ObjectPropertyValue -Object $package -PropertyName "id")
                            Requested = [string](Get-ObjectPropertyValue -Object $package -PropertyName "requestedVersion")
                            Resolved = [string](Get-ObjectPropertyValue -Object $package -PropertyName "resolvedVersion")
                            Severity = [string](Get-ObjectPropertyValue -Object $vulnerability -PropertyName "severity")
                            AdvisoryUrl = [string](Get-ObjectPropertyValue -Object $vulnerability -PropertyName "advisoryurl")
                        }
                    }
                }
            }
        }
    }

    return $records
}

$resolvedSolutionPath = Resolve-RepoPath -BaseDirectory $repositoryRoot -PathValue $SolutionPath
$resolvedReportPath = Resolve-RepoPath -BaseDirectory $repositoryRoot -PathValue $ReportPath
$reportDirectory = Split-Path -Parent $resolvedReportPath

if (-not (Test-Path $resolvedSolutionPath -PathType Leaf)) {
    Stop-WithError "Solution file not found: $resolvedSolutionPath"
}

if (-not (Test-Path $reportDirectory)) {
    New-Item -Path $reportDirectory -ItemType Directory -Force | Out-Null
}

Write-Step "Scanning NuGet dependencies for known vulnerabilities"
Write-Info "Solution : $resolvedSolutionPath"
Write-Info "Report   : $resolvedReportPath"
Write-Info "Locked restore before inspection: $([bool]$LockedMode)"

Invoke-SolutionRestoreForPackageInspection -SolutionPath $resolvedSolutionPath -UseLockedMode ([bool]$LockedMode)

$projectPaths = Get-SolutionProjectPaths -SolutionPath $resolvedSolutionPath
$projectPayloads = @()
$aggregatedSources = @()
$aggregatedProblems = @()
$records = @()

foreach ($projectPath in $projectPaths) {
    $arguments = @(
        "list", $projectPath,
        "package",
        "--vulnerable",
        "--format", "json"
    )

    if ($IncludeTransitive) {
        $arguments += "--include-transitive"
    }

    $rawOutput = & dotnet @arguments
    if ($LASTEXITCODE -ne 0) {
        Stop-WithError "dotnet list package --vulnerable failed for project $projectPath (exit code: $LASTEXITCODE)" $LASTEXITCODE
    }

    if ([string]::IsNullOrWhiteSpace($rawOutput)) {
        Stop-WithError "dotnet list package --vulnerable returned no output for project $projectPath."
    }

    try {
        $payload = $rawOutput | ConvertFrom-Json -Depth 100
    }
    catch {
        Stop-WithError "Unable to parse vulnerability scan output as JSON for project $projectPath. $($_.Exception.Message)"
    }

    $aggregatedSources += @(
        @(Get-ObjectPropertyValue -Object $payload -PropertyName "sources") |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    $aggregatedProblems += @(
        @(Get-ObjectPropertyValue -Object $payload -PropertyName "problems") |
            Where-Object { $null -ne $_ }
    )

    $records += @(Get-VulnerabilityRecords -Payload $payload)
    $projectPayloads += $payload
}

$reportProjects = foreach ($payload in $projectPayloads) {
    foreach ($project in @(Get-ObjectPropertyValue -Object $payload -PropertyName "projects")) {
        if ($null -eq $project) {
            continue
        }

        [ordered]@{
            path = [string](Get-ObjectPropertyValue -Object $project -PropertyName "path")
            frameworks = @(
                foreach ($framework in @(Get-ObjectPropertyValue -Object $project -PropertyName "frameworks")) {
                    if ($null -eq $framework) {
                        continue
                    }

                    [ordered]@{
                        framework = [string](Get-ObjectPropertyValue -Object $framework -PropertyName "framework")
                    }
                }
            )
        }
    }
}

$reportProblems = foreach ($problem in $aggregatedProblems) {
    [ordered]@{
        project = [string](Get-ObjectPropertyValue -Object $problem -PropertyName "project")
        level = [string](Get-ObjectPropertyValue -Object $problem -PropertyName "level")
        text = [string](Get-ObjectPropertyValue -Object $problem -PropertyName "text")
    }
}

$reportPayload = [ordered]@{
    generated_at_utc = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
    solution = $resolvedSolutionPath
    parameters = if ($IncludeTransitive) { "--vulnerable --include-transitive" } else { "--vulnerable" }
    restore_locked_mode = [bool]$LockedMode
    sources = @($aggregatedSources | Sort-Object -Unique)
    problems = @($reportProblems)
    projects = @($reportProjects)
}

$reportPayload | ConvertTo-Json -Depth 100 | Out-File -FilePath $resolvedReportPath -Encoding utf8

if ($records.Count -gt 0) {
    foreach ($record in $records) {
        Write-Host ("   [VULNERABLE] {0} | {1} | {2} -> {3} | severity={4} | {5}" -f `
                $record.Project,
                $record.Framework,
                $record.PackageId,
                $record.Resolved,
                $record.Severity,
                $record.AdvisoryUrl) -ForegroundColor Red
    }

    Stop-WithError "NuGet vulnerability scan found $($records.Count) vulnerable package reference(s). See $resolvedReportPath."
}

if ($aggregatedProblems.Count -gt 0) {
    foreach ($problem in $aggregatedProblems) {
        Write-Host ("   [PROBLEM] {0} | {1} | {2}" -f `
                (Get-ObjectPropertyValue -Object $problem -PropertyName "project"),
                (Get-ObjectPropertyValue -Object $problem -PropertyName "level"),
                (Get-ObjectPropertyValue -Object $problem -PropertyName "text")) -ForegroundColor Yellow
    }

    Stop-WithError "NuGet vulnerability scan completed with $($aggregatedProblems.Count) tooling problem(s). See $resolvedReportPath."
}

Write-Success "No known NuGet vulnerabilities reported across $($projectPaths.Count) project(s)."
