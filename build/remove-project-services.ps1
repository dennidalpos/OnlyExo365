[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string[]]$NameFragment = @(
        "ExchangeAdmin",
        "OnlyExo365"
    ),

    [string[]]$PathHint = @()
)

$ErrorActionPreference = "Stop"

function Write-Info {
    param([string]$Message)

    Write-Host "[INFO] $Message" -ForegroundColor Cyan
}

function Write-Success {
    param([string]$Message)

    Write-Host "[OK] $Message" -ForegroundColor Green
}

function Write-Warn {
    param([string]$Message)

    Write-Warning $Message
}

function Normalize-Text {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    return $Value.Trim()
}

function Normalize-PathHint {
    param([string]$Value)

    $normalized = Normalize-Text -Value $Value
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $null
    }

    try {
        return [System.IO.Path]::GetFullPath($normalized).TrimEnd('\')
    }
    catch {
        return $normalized.TrimEnd('\')
    }
}

function Test-ContainsAnyFragment {
    param(
        [string]$Candidate,
        [string[]]$Fragments
    )

    $normalizedCandidate = Normalize-Text -Value $Candidate
    if ([string]::IsNullOrWhiteSpace($normalizedCandidate)) {
        return $false
    }

    foreach ($fragment in $Fragments) {
        $normalizedFragment = Normalize-Text -Value $fragment
        if ([string]::IsNullOrWhiteSpace($normalizedFragment)) {
            continue
        }

        if ($normalizedCandidate.IndexOf($normalizedFragment, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
            return $true
        }
    }

    return $false
}

function Test-MatchesPathHints {
    param(
        [string]$Candidate,
        [string[]]$Hints
    )

    $normalizedCandidate = Normalize-Text -Value $Candidate
    if ([string]::IsNullOrWhiteSpace($normalizedCandidate)) {
        return $false
    }

    foreach ($hint in $Hints) {
        if ([string]::IsNullOrWhiteSpace($hint)) {
            continue
        }

        if ($normalizedCandidate.StartsWith($hint, [System.StringComparison]::OrdinalIgnoreCase) -or
            $normalizedCandidate.IndexOf($hint, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
            return $true
        }
    }

    return $false
}

function Get-ServiceParameters {
    param([string]$ServiceName)

    $serviceParametersPath = "HKLM:\SYSTEM\CurrentControlSet\Services\$ServiceName\Parameters"
    try {
        return Get-ItemProperty -Path $serviceParametersPath -ErrorAction Stop
    }
    catch {
        return $null
    }
}

function Get-ProjectServiceCandidates {
    param(
        [string[]]$NameFragments,
        [string[]]$NormalizedPathHints
    )

    $services = @(Get-CimInstance Win32_Service -ErrorAction SilentlyContinue)
    $candidates = New-Object System.Collections.Generic.List[object]

    foreach ($service in $services) {
        if ($null -eq $service) {
            continue
        }

        $parameters = Get-ServiceParameters -ServiceName $service.Name
        $searchValues = @(
            $service.Name,
            $service.DisplayName,
            $service.PathName,
            $parameters.Application,
            $parameters.AppDirectory,
            $parameters.ImagePath
        )

        $matchesFragments = $false
        foreach ($value in $searchValues) {
            if (Test-ContainsAnyFragment -Candidate $value -Fragments $NameFragments) {
                $matchesFragments = $true
                break
            }
        }

        $matchesPaths = $false
        foreach ($value in $searchValues) {
            if (Test-MatchesPathHints -Candidate $value -Hints $NormalizedPathHints) {
                $matchesPaths = $true
                break
            }
        }

        if (-not $matchesFragments -and -not $matchesPaths) {
            continue
        }

        $candidates.Add([pscustomobject]@{
                Name = $service.Name
                DisplayName = $service.DisplayName
                State = $service.State
                StartMode = $service.StartMode
                PathName = $service.PathName
                ProcessId = $service.ProcessId
                Application = $parameters.Application
                AppDirectory = $parameters.AppDirectory
            })
    }

    return @($candidates | Sort-Object Name -Unique)
}

function Stop-ServiceRobust {
    param([string]$ServiceName)

    try {
        $service = Get-Service -Name $ServiceName -ErrorAction Stop
    }
    catch {
        return
    }

    if ($service.Status -eq [System.ServiceProcess.ServiceControllerStatus]::Stopped) {
        return
    }

    try {
        Stop-Service -Name $ServiceName -Force -ErrorAction Stop
    }
    catch {
        Write-Warn "Stop-Service failed for '$ServiceName': $($_.Exception.Message). Retrying with sc.exe."
        & sc.exe stop $ServiceName | Out-Null
    }

    $deadline = (Get-Date).AddSeconds(20)
    do {
        Start-Sleep -Milliseconds 500
        try {
            $service.Refresh()
        }
        catch {
            return
        }
    } while ($service.Status -ne [System.ServiceProcess.ServiceControllerStatus]::Stopped -and (Get-Date) -lt $deadline)
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " ExchangeAdmin Service Cleanup" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$normalizedPathHints = @(
    $PathHint |
    ForEach-Object { Normalize-PathHint -Value $_ } |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
    Select-Object -Unique
)

$candidates = @(Get-ProjectServiceCandidates -NameFragments $NameFragment -NormalizedPathHints $normalizedPathHints)

if ($candidates.Count -eq 0) {
    Write-Success "No ExchangeAdmin / OnlyExo365 services found."
    exit 0
}

Write-Info "Found $($candidates.Count) candidate service(s)."

foreach ($candidate in $candidates) {
    $description = if ([string]::IsNullOrWhiteSpace($candidate.DisplayName)) {
        $candidate.Name
    }
    else {
        "$($candidate.Name) ($($candidate.DisplayName))"
    }

    Write-Info "Candidate: $description"
    Write-Info "Path: $($candidate.PathName)"

    if (-not $PSCmdlet.ShouldProcess($candidate.Name, "Stop and delete service")) {
        continue
    }

    Stop-ServiceRobust -ServiceName $candidate.Name

    if ($candidate.ProcessId -gt 0) {
        try {
            Stop-Process -Id $candidate.ProcessId -Force -ErrorAction SilentlyContinue
        }
        catch {
        }
    }

    & sc.exe delete $candidate.Name | Out-Null

    $deadline = (Get-Date).AddSeconds(20)
    $removed = $false
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 500
        $remaining = Get-CimInstance Win32_Service -Filter ("Name='{0}'" -f $candidate.Name.Replace("'", "''")) -ErrorAction SilentlyContinue
        if ($null -eq $remaining) {
            $removed = $true
            break
        }
    }

    if ($removed) {
        Write-Success "Removed service '$($candidate.Name)'."
    }
    else {
        Write-Warn "Service '$($candidate.Name)' is pending deletion or requires elevated rights."
    }
}

Write-Host ""
exit 0
