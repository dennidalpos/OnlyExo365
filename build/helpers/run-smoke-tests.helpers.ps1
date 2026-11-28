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
    $frameworksProperty = $runtimeConfig.runtimeOptions.PSObject.Properties["frameworks"]
    if ($null -ne $frameworksProperty -and $null -ne $frameworksProperty.Value) {
        foreach ($framework in @($frameworksProperty.Value)) {
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
    }

    $frameworkProperty = $runtimeConfig.runtimeOptions.PSObject.Properties["framework"]
    if ($null -ne $frameworkProperty -and $null -ne $frameworkProperty.Value) {
        $frameworkEntries.Add([pscustomobject]@{
                Component = $ComponentName
                RuntimeConfigPath = $RuntimeConfigPath
                Framework = [string]$frameworkProperty.Value.name
                MinimumVersion = [string]$frameworkProperty.Value.version
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
            $actualDisplayName = [string](Get-ObjectPropertyValue -Object $_ -PropertyName "DisplayName")
            $actualPublisher = [string](Get-ObjectPropertyValue -Object $_ -PropertyName "Publisher")
            $displayNameMatches = -not [string]::IsNullOrWhiteSpace($actualDisplayName) -and (
                $actualDisplayName -eq $DisplayName -or
                $actualDisplayName.StartsWith("$DisplayName ", [System.StringComparison]::OrdinalIgnoreCase))

            $displayNameMatches -and
            ([string]::IsNullOrWhiteSpace($Publisher) -or $actualPublisher -eq $Publisher)
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
            (Normalize-DirectoryPath -PathValue ([string](Get-ObjectPropertyValue -Object $_ -PropertyName "InstallLocation"))),
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
                (Normalize-DirectoryPath -PathValue ([string](Get-ObjectPropertyValue -Object $_ -PropertyName "InstallLocation"))),
                $expectedInstallLocation,
                [System.StringComparison]::OrdinalIgnoreCase)
        })
}

function Resolve-SmokePublishPath {
    param(
        [string]$PublishRootPath
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

function Assert-InstalledPayload {
    param([string]$PublishRoot, [string]$InstallRoot, [string]$SetupPath)
    foreach ($name in @('OnlyExo365.Shell.exe', 'OnlyExo365.Worker.exe')) {
        $published = Get-FileHash -LiteralPath (Join-Path $PublishRoot $name) -Algorithm SHA256
        $installed = Get-FileHash -LiteralPath (Join-Path $InstallRoot $name) -Algorithm SHA256
        if ($published.Hash -ne $installed.Hash) { throw "Installed payload differs from publish: $name" }
    }
    $setupSignature = Get-AuthenticodeSignature -LiteralPath $SetupPath
    if ($setupSignature.Status -eq 'NotSigned') { return $false }
    if ($setupSignature.Status -ne 'Valid') { throw 'Setup Authenticode signature is invalid.' }
    foreach ($file in Get-ChildItem -LiteralPath $InstallRoot -Recurse -File -Filter '*.exe') {
        if ($file.Name -eq 'unins000.exe') { continue }
        $signature = Get-AuthenticodeSignature -LiteralPath $file.FullName
        if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Thumbprint -ne $setupSignature.SignerCertificate.Thumbprint) {
            throw "Installed payload signature does not match setup signer: $($file.FullName)"
        }
    }
    return $true
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
            $matchingEntries = @($entries | Where-Object {
                    $_.ProcessId -eq $presentationProcess.Id -and
                    [datetime]$_.TimestampUtc -ge $launchUtc.AddSeconds(-5)
                })

            if ($matchingEntries.Count -gt 0) {
                return $matchingEntries
            }

            return $null
        }

        $supervisorLogEntries = Wait-Until -TimeoutSeconds $LogProbeTimeoutSeconds -FailureMessage "$Label supervisor log entries were not found." -Condition {
            $entries = Read-JsonLinesFromCandidatePaths -LogDirectory $LogDirectory -Component "supervisor"
            $matchingEntries = @($entries | Where-Object {
                    [datetime]$_.TimestampUtc -ge $launchUtc.AddSeconds(-5)
                })

            if ($matchingEntries.Count -gt 0) {
                return $matchingEntries
            }

            return $null
        }

        $workerLogEntries = Wait-Until -TimeoutSeconds $LogProbeTimeoutSeconds -FailureMessage "$Label worker log entries were not found." -Condition {
            $entries = Read-JsonLinesFromCandidatePaths -LogDirectory $LogDirectory -Component "worker"
            $matchingEntries = @($entries | Where-Object {
                    $_.ProcessId -eq $workerInfo.ProcessId -and
                    [datetime]$_.TimestampUtc -ge $launchUtc.AddSeconds(-5)
                })

            if ($matchingEntries.Count -gt 0) {
                return $matchingEntries
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
        try {
            Stop-ProcessIfRunning -Process $presentationProcess
        }
        finally {
            Stop-ProcessIfRunning -Process $workerProcess
        }
    }
}

function Get-ProductDataHashes {
    param([string]$RootPath)
    if ((Get-Item -LiteralPath $RootPath -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Recovery refuses reparse point: $RootPath" }
    $hashes = [ordered]@{}
    $pending = [Collections.Generic.Stack[string]]::new()
    $pending.Push($RootPath)
    while ($pending.Count) {
        foreach ($item in Get-ChildItem -LiteralPath $pending.Pop() -Force | Sort-Object Name) {
            if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Backup refuses reparse point: $($item.FullName)" }
            if ($item.PSIsContainer) { $pending.Push($item.FullName) }
            else { $hashes[[IO.Path]::GetRelativePath($RootPath, $item.FullName)] = (Get-FileHash -LiteralPath $item.FullName -Algorithm SHA256).Hash }
        }
    }
    return $hashes
}

function Save-ProductDataRecoveryManifest {
    param([hashtable]$RecoveryState)
    $temporaryPath = $RecoveryState.ManifestPath + '.tmp'
    try {
        $RecoveryState | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $temporaryPath -Encoding utf8NoBOM
        [IO.File]::Move($temporaryPath, $RecoveryState.ManifestPath, $true)
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath) { Remove-Item -LiteralPath $temporaryPath -ErrorAction Stop }
    }
}

function Backup-ProductDataRoots {
    param([string[]]$Paths, [hashtable]$RecoveryState = @{})
    $RecoveryState.BackupRoot = Join-Path ([IO.Path]::GetTempPath()) ('onlyexo365-localdata-backup-' + [guid]::NewGuid().ToString('N'))
    $RecoveryState.ManifestPath = Join-Path $RecoveryState.BackupRoot 'recovery.json'
    $RecoveryState.Entries = [Collections.Generic.List[object]]::new()
    $normalizedPaths = @($Paths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { Normalize-FileSystemPath $_ } | Sort-Object -Unique)
    foreach ($path in $normalizedPaths) {
        if (-not (Test-Path -LiteralPath $path)) { continue }
        if ((Get-Item -LiteralPath $path -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Backup refuses reparse point: $path" }
        $entry = [ordered]@{
            original_path = $path
            backup_path = Join-Path $RecoveryState.BackupRoot ([guid]::NewGuid().ToString('N'))
            content_hashes = Get-ProductDataHashes -RootPath $path
            created_data_path = $null
        }
        New-Item -Path $RecoveryState.BackupRoot -ItemType Directory -Force | Out-Null
        $RecoveryState.Entries.Add($entry)
        # Persist recovery before moving data.
        Save-ProductDataRecoveryManifest -RecoveryState $RecoveryState
        Write-Warn "Moving it to temporary backup before installer uninstall validation: $path. Recovery manifest: $($RecoveryState.ManifestPath)"
        Move-Item -LiteralPath $path -Destination $entry.backup_path -ErrorAction Stop
    }
    return $RecoveryState
}

function Restore-ProductDataRoots {
    param([object[]]$Entries, [hashtable]$RecoveryState)
    $restoredCount = 0
    foreach ($entry in @($Entries)) {
        $originalPath = Normalize-FileSystemPath $entry.original_path
        $backupPath = Normalize-FileSystemPath $entry.backup_path
        if (Test-Path -LiteralPath $backupPath) {
            $backupHashes = Get-ProductDataHashes -RootPath $backupPath
            if ($backupHashes.Count -ne $entry.content_hashes.Count) { throw "Recovery backup is incomplete: $backupPath" }
            foreach ($key in $entry.content_hashes.Keys) {
                if (-not $backupHashes.Contains($key) -or $backupHashes[$key] -ne $entry.content_hashes[$key]) { throw "Recovery backup content differs: $backupPath/$key" }
            }
            if (Test-Path -LiteralPath $originalPath) {
                $entry.created_data_path = Join-Path (Split-Path -Parent $backupPath) ('smoke-output-' + [guid]::NewGuid().ToString('N'))
                if ($null -ne $RecoveryState) { Save-ProductDataRecoveryManifest -RecoveryState $RecoveryState }
                Move-Item -LiteralPath $originalPath -Destination $entry.created_data_path -ErrorAction Stop
            }
            New-Item -Path (Split-Path -Parent $originalPath) -ItemType Directory -Force | Out-Null
            Move-Item -LiteralPath $backupPath -Destination $originalPath -ErrorAction Stop
        }
        if (-not (Test-Path -LiteralPath $originalPath)) { throw "Recovery missing original and backup: $originalPath; backup: $backupPath" }
        $actualHashes = Get-ProductDataHashes -RootPath $originalPath
        if ($actualHashes.Count -ne $entry.content_hashes.Count) { throw "Recovery content count differs: $originalPath" }
        foreach ($key in $entry.content_hashes.Keys) {
            if (-not $actualHashes.Contains($key) -or $actualHashes[$key] -ne $entry.content_hashes[$key]) { throw "Recovery content differs: $originalPath/$key" }
        }
        $restoredCount++
    }
    return $restoredCount
}

function Complete-ProductDataRecovery {
    param([hashtable]$RecoveryState)
    if (-not (Test-Path -LiteralPath $RecoveryState.BackupRoot)) { return }
    if (@($RecoveryState.Entries | Where-Object { $null -ne $_.created_data_path }).Count -gt 0) {
        Write-Info "Preserved smoke-created data and recovery manifest at $($RecoveryState.BackupRoot)"
        return
    }
    $remaining = @(Get-ChildItem -LiteralPath $RecoveryState.BackupRoot -Force | Where-Object FullName -ne $RecoveryState.ManifestPath)
    if ($remaining.Count -gt 0) { throw "Recovery retains unexpected files and manifest at $($RecoveryState.BackupRoot)" }
    Remove-Item -LiteralPath $RecoveryState.ManifestPath -ErrorAction Stop
    # Remove only an empty, recovered backup.
    [IO.Directory]::Delete($RecoveryState.BackupRoot)
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

