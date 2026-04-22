function Get-RequiredRuntimeEntries {
    param(
        [string]$ComponentName,
        [string]$RuntimeConfigPath
    )

    if (-not (Test-Path $RuntimeConfigPath -PathType Leaf)) {
        Stop-WithError "Runtime config not found: $RuntimeConfigPath"
    }

    try {
        $runtimeConfig = Get-Content -Raw $RuntimeConfigPath | ConvertFrom-Json -Depth 20
    }
    catch {
        Stop-WithError "Unable to parse runtime config: $RuntimeConfigPath"
    }

    $frameworkEntries = New-Object System.Collections.Generic.List[object]
    foreach ($framework in @($runtimeConfig.runtimeOptions.frameworks)) {
        if ($null -eq $framework) {
            continue
        }

        $frameworkEntries.Add([pscustomobject]@{
                Component = $ComponentName
                RuntimeConfigPath = $RuntimeConfigPath
                Framework = [string]$framework.name
                MinimumVersion = [string]$framework.version
            })
    }

    if ($null -ne $runtimeConfig.runtimeOptions.framework) {
        $frameworkEntries.Add([pscustomobject]@{
                Component = $ComponentName
                RuntimeConfigPath = $RuntimeConfigPath
                Framework = [string]$runtimeConfig.runtimeOptions.framework.name
                MinimumVersion = [string]$runtimeConfig.runtimeOptions.framework.version
            })
    }

    return $frameworkEntries
}

function Get-InstalledRuntimeCatalog {
    $rawLines = & dotnet --list-runtimes 2>&1
    if ($LASTEXITCODE -ne 0) {
        Stop-WithError "Unable to enumerate installed .NET runtimes via 'dotnet --list-runtimes'."
    }

    $catalog = @{}
    foreach ($line in @($rawLines)) {
        $text = [string]$line
        if ($text -match '^(?<name>\S+)\s+(?<version>\d+\.\d+\.\d+)\s+\[(?<path>.+)\]$') {
            $name = $Matches.name
            if (-not $catalog.ContainsKey($name)) {
                $catalog[$name] = New-Object System.Collections.Generic.List[string]
            }

            $catalog[$name].Add($Matches.version)
        }
    }

    return $catalog
}

function Test-RuntimeRequirementSatisfied {
    param(
        [string]$RequiredVersion,
        [string[]]$InstalledVersions
    )

    if ([string]::IsNullOrWhiteSpace($RequiredVersion)) {
        return $false
    }

    try {
        $required = [Version]$RequiredVersion
    }
    catch {
        return $false
    }

    foreach ($installedVersion in @($InstalledVersions)) {
        try {
            $candidate = [Version]$installedVersion
            if ($candidate.Major -eq $required.Major -and $candidate.Minor -ge $required.Minor -and $candidate -ge $required) {
                return $true
            }
        }
        catch {
            continue
        }
    }

    return $false
}

function Get-PowerShell7Info {
    $command = Get-Command pwsh -ErrorAction SilentlyContinue
    if ($null -eq $command -or [string]::IsNullOrWhiteSpace($command.Source)) {
        Stop-WithError "PowerShell 7 runtime not found in PATH."
    }

    $versionText = & $command.Source -NoLogo -NoProfile -Command '$PSVersionTable.PSVersion.ToString()'
    if ($LASTEXITCODE -ne 0) {
        Stop-WithError "Unable to determine PowerShell runtime version from pwsh."
    }

    try {
        $version = [Version]$versionText
    }
    catch {
        Stop-WithError "PowerShell runtime version is invalid: $versionText"
    }

    if ($version.Major -lt 7) {
        Stop-WithError "PowerShell 7 or later is required for packaged execution. Found: $versionText"
    }

    return [ordered]@{
        path = $command.Source
        version = $version.ToString()
        minimum_major = 7
        satisfies_minimum = $true
    }
}

function Get-PrerequisiteReport {
    param(
        [string]$PresentationRuntimeConfig,
        [string]$WorkerRuntimeConfig
    )

    $runtimeCatalog = Get-InstalledRuntimeCatalog
    $requirements = @()
    $requirements += @(Get-RequiredRuntimeEntries -ComponentName "OnlyExo365.Shell" -RuntimeConfigPath $PresentationRuntimeConfig)
    $requirements += @(Get-RequiredRuntimeEntries -ComponentName "OnlyExo365.Worker" -RuntimeConfigPath $WorkerRuntimeConfig)

    $runtimeEntries = New-Object System.Collections.Generic.List[object]
    foreach ($requirement in $requirements) {
        if ($null -eq $requirement) {
            continue
        }

        $installedVersions = if ($runtimeCatalog.ContainsKey($requirement.Framework)) {
            @($runtimeCatalog[$requirement.Framework].ToArray())
        }
        else {
            @()
        }

        $satisfied = Test-RuntimeRequirementSatisfied -RequiredVersion $requirement.MinimumVersion -InstalledVersions $installedVersions
        if (-not $satisfied) {
            Stop-WithError "Missing required runtime '$($requirement.Framework) >= $($requirement.MinimumVersion)' for $($requirement.Component). Installed: $($installedVersions -join ', ')"
        }

        $runtimeEntries.Add([ordered]@{
                component = $requirement.Component
                runtime_config = $requirement.RuntimeConfigPath
                framework = $requirement.Framework
                minimum_version = $requirement.MinimumVersion
                installed_versions = $installedVersions
                satisfies_minimum = $satisfied
            })
    }

    return [ordered]@{
        powershell = Get-PowerShell7Info
        runtimes = $runtimeEntries.ToArray()
    }
}

function Get-UninstallEntries {
    param(
        [string]$DisplayName,
        [string]$Publisher
    )

    $registryPaths = @(
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    return @(Get-ItemProperty -Path $registryPaths -ErrorAction SilentlyContinue | Where-Object {
            $actualDisplayName = [string]$_.DisplayName
            $displayNameMatches = -not [string]::IsNullOrWhiteSpace($actualDisplayName) -and (
                $actualDisplayName -eq $DisplayName -or
                $actualDisplayName.StartsWith("$DisplayName ", [System.StringComparison]::OrdinalIgnoreCase))

            $displayNameMatches -and
            ([string]::IsNullOrWhiteSpace($Publisher) -or $_.Publisher -eq $Publisher)
        })
}

function Normalize-DirectoryPath {
    param([string]$PathValue)

    if ([string]::IsNullOrWhiteSpace($PathValue)) {
        return $PathValue
    }

    return $PathValue.TrimEnd('\')
}

function Get-InstalledProductEntry {
    param(
        [string]$DisplayName,
        [string]$Publisher,
        [string]$InstallLocation
    )

    $entries = @(Get-UninstallEntries -DisplayName $DisplayName -Publisher $Publisher)
    if ([string]::IsNullOrWhiteSpace($InstallLocation)) {
        return $entries | Select-Object -First 1
    }

    $expectedInstallLocation = Normalize-DirectoryPath -PathValue $InstallLocation
    $locationMatch = $entries | Where-Object {
        [string]::Equals(
            (Normalize-DirectoryPath -PathValue $_.InstallLocation),
            $expectedInstallLocation,
            [System.StringComparison]::OrdinalIgnoreCase)
    } | Select-Object -First 1

    if ($null -ne $locationMatch) {
        return $locationMatch
    }

    return $null
}

function Get-InstalledProductEntriesForLocation {
    param(
        [string]$DisplayName,
        [string]$Publisher,
        [string]$InstallLocation
    )

    $entries = @(Get-UninstallEntries -DisplayName $DisplayName -Publisher $Publisher)
    if ([string]::IsNullOrWhiteSpace($InstallLocation)) {
        return $entries
    }

    $expectedInstallLocation = Normalize-DirectoryPath -PathValue $InstallLocation
    return @($entries | Where-Object {
            [string]::Equals(
                (Normalize-DirectoryPath -PathValue $_.InstallLocation),
                $expectedInstallLocation,
                [System.StringComparison]::OrdinalIgnoreCase)
        })
}

function Resolve-SmokePublishPath {
    param(
        [string]$PublishRootPath,
        [bool]$Prefer64BitPayload
    )

    $presentationExeAtRoot = Join-Path $PublishRootPath "OnlyExo365.Shell.exe"
    if (Test-Path $presentationExeAtRoot -PathType Leaf) {
        return $PublishRootPath
    }

    $runtimeIdentifier = "win-x64"
    $runtimePublishPath = Join-Path $PublishRootPath $runtimeIdentifier
    if (Test-Path $runtimePublishPath -PathType Container) {
        return $runtimePublishPath
    }

    Stop-WithError "Publish output for runtime '$runtimeIdentifier' not found under $PublishRootPath"
}

function Get-UninstallInvocation {
    param([string]$CommandLine)

    if ([string]::IsNullOrWhiteSpace($CommandLine)) {
        Stop-WithError "Uninstall command line is empty."
    }

    if ($CommandLine -match '^\s*"(?<file>[^"]+)"\s*(?<args>.*)$') {
        return [ordered]@{
            file_path = $Matches.file
            arguments = $Matches.args
        }
    }

    if ($CommandLine -match '^\s*(?<file>\S+)\s*(?<args>.*)$') {
        return [ordered]@{
            file_path = $Matches.file
            arguments = $Matches.args
        }
    }

    Stop-WithError "Unable to parse uninstall command line: $CommandLine"
}

function Get-RequiredPayloadFiles {
    param([string]$RootPath)

    return @(
        (Join-Path $RootPath "OnlyExo365.Shell.exe"),
        (Join-Path $RootPath "OnlyExo365.Worker.exe"),
        (Join-Path $RootPath "appsettings.json")
    )
}

function Assert-RequiredFilesExist {
    param(
        [string[]]$Paths,
        [string]$Label
    )

    foreach ($path in @($Paths)) {
        if (-not (Test-Path $path -PathType Leaf)) {
            Stop-WithError "$Label missing expected file: $path"
        }
    }
}

function Invoke-ApplicationSmokeTest {
    param(
        [string]$Label,
        [string]$ExecutablePath,
        [string]$WorkerExecutablePath,
        [string]$WorkingDirectory,
        [string]$LogDirectory,
        [int]$StartupTimeoutSeconds,
        [int]$StabilityWindowSeconds,
        [int]$LogProbeTimeoutSeconds
    )

    $presentationProcess = $null
    $workerProcess = $null
    $launchUtc = (Get-Date).ToUniversalTime()

    try {
        Write-Step "Running $Label smoke test"
        Write-Info "Executable: $ExecutablePath"

        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = $ExecutablePath
        $startInfo.WorkingDirectory = $WorkingDirectory
        $startInfo.UseShellExecute = $false
        $startInfo.EnvironmentVariables["ONLYEXO365_DISABLE_EXO"] = "1"

        $presentationProcess = [System.Diagnostics.Process]::Start($startInfo)
        if ($null -eq $presentationProcess) {
            Stop-WithError "Unable to start $Label executable."
        }

        Write-Info "Presentation PID: $($presentationProcess.Id)"

        Wait-Until -TimeoutSeconds $StartupTimeoutSeconds -FailureMessage "$Label presentation process exited during startup." -Condition {
            try {
                if ($presentationProcess.HasExited) {
                    return $null
                }

                return $presentationProcess
            }
            catch {
                return $null
            }
        } | Out-Null

        $workerInfo = Wait-Until -TimeoutSeconds $StartupTimeoutSeconds -FailureMessage "$Label worker process did not appear." -Condition {
            Get-WorkerProcessInfo -ExpectedPath $WorkerExecutablePath -LaunchedAfterUtc $launchUtc
        }

        Write-Info "Worker PID: $($workerInfo.ProcessId)"
        $workerProcess = Get-Process -Id $workerInfo.ProcessId -ErrorAction Stop

        Write-Info "Observing process stability for $StabilityWindowSeconds seconds"
        $stableUntil = (Get-Date).AddSeconds($StabilityWindowSeconds)
        while ((Get-Date) -lt $stableUntil) {
            if ($presentationProcess.HasExited) {
                Stop-WithError "$Label presentation process exited before the stability window completed."
            }

            try {
                $workerProcess.Refresh()
                if ($workerProcess.HasExited) {
                    Stop-WithError "$Label worker process exited before the stability window completed."
                }
            }
            catch {
                Stop-WithError "$Label worker process disappeared before the stability window completed."
            }

            Start-Sleep -Milliseconds 500
        }

        Write-Success "$Label process stability confirmed."

        Write-Step "Validating $Label logs"
        Write-Info "Log directory: $LogDirectory"

        $uiLogEntries = Wait-Until -TimeoutSeconds $LogProbeTimeoutSeconds -FailureMessage "$Label UI log entries for the launched process were not found." -Condition {
            $entries = Read-JsonLinesFromCandidatePaths -LogDirectory $LogDirectory -Component "ui"
            $matches = @($entries | Where-Object {
                    $_.ProcessId -eq $presentationProcess.Id -and
                    [datetime]$_.TimestampUtc -ge $launchUtc.AddSeconds(-5)
                })

            if ($matches.Count -gt 0) {
                return $matches
            }

            return $null
        }

        $supervisorLogEntries = Wait-Until -TimeoutSeconds $LogProbeTimeoutSeconds -FailureMessage "$Label supervisor log entries were not found." -Condition {
            $entries = Read-JsonLinesFromCandidatePaths -LogDirectory $LogDirectory -Component "supervisor"
            $matches = @($entries | Where-Object {
                    [datetime]$_.TimestampUtc -ge $launchUtc.AddSeconds(-5)
                })

            if ($matches.Count -gt 0) {
                return $matches
            }

            return $null
        }

        $workerLogEntries = Wait-Until -TimeoutSeconds $LogProbeTimeoutSeconds -FailureMessage "$Label worker log entries were not found." -Condition {
            $entries = Read-JsonLinesFromCandidatePaths -LogDirectory $LogDirectory -Component "worker"
            $matches = @($entries | Where-Object {
                    $_.ProcessId -eq $workerInfo.ProcessId -and
                    [datetime]$_.TimestampUtc -ge $launchUtc.AddSeconds(-5)
                })

            if ($matches.Count -gt 0) {
                return $matches
            }

            return $null
        }

        if (-not (@($uiLogEntries | Where-Object { $_.Message -like "*Worker state changed: Connected*" }).Count -gt 0)) {
            Stop-WithError "$Label UI log does not confirm worker connection for the launched process."
        }

        if (-not (@($supervisorLogEntries | Where-Object { $_.Message -like "*IPC connection successful*" -or $_.Message -like "*Connected. Module available*" }).Count -gt 0)) {
            Stop-WithError "$Label supervisor log does not confirm worker handshake."
        }

        if (-not (@($workerLogEntries | Where-Object { $_.Message -like "*IPC server started. Waiting for connections...*" }).Count -gt 0)) {
            Stop-WithError "$Label worker log does not confirm IPC server startup."
        }

        Write-Success "$Label logs confirm UI bootstrap, worker launch and IPC startup."

        return [ordered]@{
            presentation_process_id = $presentationProcess.Id
            worker_process_id = $workerInfo.ProcessId
            ui_connected_log_confirmed = $true
            supervisor_log_confirmed = $true
            worker_log_confirmed = $true
            launched_at_utc = $launchUtc.ToString("O")
        }
    }
    finally {
        Stop-ProcessIfRunning -Process $presentationProcess

        if ($null -ne $workerProcess) {
            try {
                Stop-Process -Id $workerProcess.Id -Force -ErrorAction SilentlyContinue
            }
            catch {
            }
        }
    }
}

function Get-ProjectServiceEntries {
    param([string[]]$PathHints)

    $normalizedHints = @(
        $PathHints |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { Normalize-FileSystemPath -PathValue $_ } |
        Select-Object -Unique
    )

    $services = @(Get-CimInstance Win32_Service -ErrorAction SilentlyContinue)
    $matches = New-Object System.Collections.Generic.List[object]

    foreach ($service in $services) {
        $parametersPath = "HKLM:\SYSTEM\CurrentControlSet\Services\$($service.Name)\Parameters"
        $parameters = Get-ItemProperty -Path $parametersPath -ErrorAction SilentlyContinue

        $searchValues = @(
            $service.Name,
            $service.DisplayName,
            $service.PathName,
            $parameters.Application,
            $parameters.AppDirectory
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        $isProjectService = $false
        foreach ($value in $searchValues) {
            if ($value -match 'OnlyExo365|OnlyExo365') {
                $isProjectService = $true
                break
            }

            foreach ($hint in $normalizedHints) {
                if ($value.IndexOf($hint, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    $isProjectService = $true
                    break
                }
            }

            if ($isProjectService) {
                break
            }
        }

        if (-not $isProjectService) {
            continue
        }

        $matches.Add([ordered]@{
                name = $service.Name
                display_name = $service.DisplayName
                state = $service.State
                path = $service.PathName
            })
    }

    return @($matches.ToArray())
}

function Backup-ProjectDataRoots {
    param([string[]]$Paths)

    $normalizedPaths = @(
        $Paths |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { Normalize-FileSystemPath -PathValue $_ } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique
    )

    if ($normalizedPaths.Count -eq 0) {
        return [ordered]@{
            BackupRoot = $null
            Entries = @()
        }
    }

    $backupRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("onlyexo365-localdata-backup-" + [guid]::NewGuid().ToString("N"))
    $entries = New-Object System.Collections.Generic.List[object]

    foreach ($path in $normalizedPaths) {
        if (-not (Test-Path $path)) {
            continue
        }

        if (-not (Test-Path $backupRoot)) {
            New-Item -Path $backupRoot -ItemType Directory -Force | Out-Null
        }

        $backupPath = Join-Path $backupRoot ([guid]::NewGuid().ToString("N"))
        $backupParent = Split-Path -Parent $backupPath
        if (-not [string]::IsNullOrWhiteSpace($backupParent)) {
            New-Item -Path $backupParent -ItemType Directory -Force | Out-Null
        }

        Write-Warn "Existing OnlyExo365 local data detected at '$path'. Moving it to temporary backup before installer uninstall validation."
        Move-Item -Path $path -Destination $backupPath -Force
        $entries.Add([ordered]@{
                original_path = $path
                backup_path = $backupPath
            })
    }

    $resolvedBackupRoot = $null
    if ($entries.Count -gt 0) {
        $resolvedBackupRoot = $backupRoot
    }

    return [ordered]@{
        BackupRoot = $resolvedBackupRoot
        Entries = @($entries.ToArray())
    }
}

function Restore-ProjectDataRoots {
    param([object[]]$Entries)

    $restoredCount = 0

    foreach ($entry in @($Entries)) {
        if ($null -eq $entry) {
            continue
        }

        $originalPath = Normalize-FileSystemPath -PathValue $entry.original_path
        $backupPath = Normalize-FileSystemPath -PathValue $entry.backup_path

        if ([string]::IsNullOrWhiteSpace($originalPath) -or [string]::IsNullOrWhiteSpace($backupPath)) {
            continue
        }

        if (-not (Test-Path $backupPath)) {
            continue
        }

        if (Test-Path $originalPath) {
            Remove-Item -Path $originalPath -Recurse -Force -ErrorAction Stop
        }

        $originalParent = Split-Path -Parent $originalPath
        if (-not [string]::IsNullOrWhiteSpace($originalParent) -and -not (Test-Path $originalParent)) {
            New-Item -Path $originalParent -ItemType Directory -Force | Out-Null
        }

        Move-Item -Path $backupPath -Destination $originalPath -Force
        $restoredCount++
    }

    return $restoredCount
}

function New-ResidualMarkerFile {
    param(
        [string]$RootPath,
        [string]$RelativePath
    )

    $targetPath = Join-Path $RootPath $RelativePath
    $targetDirectory = Split-Path -Parent $targetPath
    if (-not [string]::IsNullOrWhiteSpace($targetDirectory)) {
        New-Item -Path $targetDirectory -ItemType Directory -Force | Out-Null
    }

    Set-Content -Path $targetPath -Value "residual-marker" -Encoding utf8
    return $targetPath
}

