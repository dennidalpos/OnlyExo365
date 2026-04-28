#Requires -Version 7.0

[CmdletBinding()]
param(
    [string]$PublishPath = "artifacts/publish",
    [string]$SetupExePath = "artifacts/packages/OnlyExo365.Setup.exe",
    [string]$ReportPath = "artifacts/smoke/smoke-report.json",
    [int]$StartupTimeoutSeconds = 45,
    [int]$StabilityWindowSeconds = 15,
    [int]$LogProbeTimeoutSeconds = 20,
    [int]$InstallerOperationTimeoutSeconds = 120,
    [switch]$SkipRealInstall = $false
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$repositoryRoot = Split-Path -Parent $scriptDirectory

. (Join-Path $repositoryRoot "scripts\internal\common.ps1")

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

function Normalize-FileSystemPath {
    param([string]$PathValue)

    if ([string]::IsNullOrWhiteSpace($PathValue)) {
        return $PathValue
    }

    return [System.IO.Path]::GetFullPath($PathValue)
}

function Wait-Until {
    param(
        [scriptblock]$Condition,
        [int]$TimeoutSeconds,
        [int]$PollMilliseconds = 500,
        [string]$FailureMessage
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        $result = & $Condition
        if ($null -ne $result) {
            return $result
        }

        Start-Sleep -Milliseconds $PollMilliseconds
    }

    Stop-WithError $FailureMessage
}

function Get-DefaultLogDirectory {
    return Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) "OnlyExo365\logs"
}

function Get-DefaultSecretDirectory {
    return Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) "OnlyExo365\ipc-secrets"
}

function Get-DefaultExportDirectory {
    return Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) "OnlyExo365\exports"
}

function Get-LogFilePath {
    param(
        [string]$LogDirectory,
        [string]$Component
    )

    return Join-Path $LogDirectory ("{0}-{1:yyyyMMdd}.log" -f $Component, (Get-Date))
}

function Get-LogFileCandidatePaths {
    param(
        [string]$LogDirectory,
        [string]$Component
    )

    $candidates = New-Object System.Collections.Generic.List[string]

    foreach ($timestamp in @((Get-Date), (Get-Date).ToUniversalTime())) {
        $candidate = Join-Path $LogDirectory ("{0}-{1:yyyyMMdd}.log" -f $Component, $timestamp)
        if (Test-Path $candidate -PathType Leaf) {
            [void]$candidates.Add($candidate)
        }
    }

    if (Test-Path $LogDirectory -PathType Container) {
        $latestMatches = @(Get-ChildItem -Path $LogDirectory -Filter "$Component-*.log" -File -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTimeUtc -Descending |
                Select-Object -First 2 -ExpandProperty FullName)
        foreach ($match in $latestMatches) {
            if (-not [string]::IsNullOrWhiteSpace($match) -and (Test-Path $match -PathType Leaf)) {
                [void]$candidates.Add($match)
            }
        }
    }

    return @($candidates | Select-Object -Unique)
}

function Read-JsonLines {
    param([string]$Path)

    if (-not (Test-Path $Path -PathType Leaf)) {
        return @()
    }

    $entries = New-Object System.Collections.Generic.List[object]
    foreach ($line in [System.IO.File]::ReadLines($Path)) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        try {
            $entries.Add(($line | ConvertFrom-Json -Depth 20))
        }
        catch {
            continue
        }
    }

    return $entries
}

function Read-JsonLinesFromCandidatePaths {
    param(
        [string]$LogDirectory,
        [string]$Component
    )

    $entries = New-Object System.Collections.Generic.List[object]
    foreach ($path in @(Get-LogFileCandidatePaths -LogDirectory $LogDirectory -Component $Component)) {
        foreach ($entry in @(Read-JsonLines -Path $path)) {
            if ($null -ne $entry) {
                [void]$entries.Add($entry)
            }
        }
    }

    return @($entries.ToArray())
}

function Get-WorkerProcessInfo {
    param(
        [string]$ExpectedPath,
        [datetime]$LaunchedAfterUtc
    )

    $normalizedExpectedPath = Normalize-FileSystemPath -PathValue $ExpectedPath
    $workers = Get-CimInstance Win32_Process -Filter "Name = 'OnlyExo365.Worker.exe'" -ErrorAction SilentlyContinue
    foreach ($worker in @($workers)) {
        if ($null -eq $worker -or [string]::IsNullOrWhiteSpace($worker.ExecutablePath)) {
            continue
        }

        $normalizedWorkerPath = Normalize-FileSystemPath -PathValue $worker.ExecutablePath
        if (-not [string]::Equals($normalizedWorkerPath, $normalizedExpectedPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }

        try {
            $process = Get-Process -Id $worker.ProcessId -ErrorAction Stop
            if ($process.StartTime.ToUniversalTime() -lt $LaunchedAfterUtc.AddSeconds(-5)) {
                continue
            }

            return [pscustomobject]@{
                ProcessId = $worker.ProcessId
                Path = $worker.ExecutablePath
                StartTimeUtc = $process.StartTime.ToUniversalTime()
            }
        }
        catch {
            continue
        }
    }

    return $null
}

function Stop-ProcessIfRunning {
    param([System.Diagnostics.Process]$Process)

    if ($null -eq $Process) {
        return
    }

    try {
        if ($Process.HasExited) {
            return
        }
    }
    catch {
        return
    }

    try {
        $null = $Process.CloseMainWindow()
        if ($Process.WaitForExit(5000)) {
            return
        }
    }
    catch {
    }

    try {
        Stop-Process -Id $Process.Id -Force -ErrorAction SilentlyContinue
    }
    catch {
    }
}

$smokeHelperPath = Join-Path $scriptDirectory "helpers\run-smoke-tests.helpers.ps1"
if (-not (Test-Path $smokeHelperPath -PathType Leaf)) {
    Stop-WithError "Smoke test helper script not found: $smokeHelperPath"
}

. $smokeHelperPath

$resolvedPublishPath = Normalize-FileSystemPath -PathValue (Resolve-RepoPath -BaseDirectory $repositoryRoot -PathValue $PublishPath)
$resolvedSetupExePath = Normalize-FileSystemPath -PathValue (Resolve-RepoPath -BaseDirectory $repositoryRoot -PathValue $SetupExePath)
$resolvedReportPath = Normalize-FileSystemPath -PathValue (Resolve-RepoPath -BaseDirectory $repositoryRoot -PathValue $ReportPath)
$resolvedPublishPath = Resolve-SmokePublishPath -PublishRootPath $resolvedPublishPath -Prefer64BitPayload ([Environment]::Is64BitOperatingSystem)
$reportDirectory = Split-Path -Parent $resolvedReportPath
$logDirectory = Get-DefaultLogDirectory
$secretDirectory = Get-DefaultSecretDirectory
$exportDirectory = Get-DefaultExportDirectory

if (-not (Test-Path $resolvedPublishPath -PathType Container)) {
    Stop-WithError "Publish path not found: $resolvedPublishPath"
}

if (-not (Test-Path $resolvedSetupExePath -PathType Leaf)) {
    Stop-WithError "Setup EXE path not found: $resolvedSetupExePath"
}

if (-not (Test-Path $reportDirectory)) {
    New-Item -Path $reportDirectory -ItemType Directory -Force | Out-Null
}

$presentationExe = Join-Path $resolvedPublishPath "OnlyExo365.Shell.exe"
$workerExe = Join-Path $resolvedPublishPath "OnlyExo365.Worker.exe"
$appSettingsPath = Join-Path $resolvedPublishPath "appsettings.json"
$presentationRuntimeConfig = Join-Path $resolvedPublishPath "OnlyExo365.Shell.runtimeconfig.json"
$workerRuntimeConfig = Join-Path $resolvedPublishPath "OnlyExo365.Worker.runtimeconfig.json"
$requiredPublishFiles = @($presentationExe, $workerExe, $appSettingsPath, $presentationRuntimeConfig, $workerRuntimeConfig)
Assert-RequiredFilesExist -Paths $requiredPublishFiles -Label "Publish output"

$setupVersion = (Get-Item -LiteralPath $resolvedSetupExePath).VersionInfo.ProductVersion
if ([string]::IsNullOrWhiteSpace($setupVersion)) {
    $setupVersion = "0.0.0"
}

$report = [ordered]@{
    generated_at_utc = (Get-Date).ToUniversalTime().ToString("O")
    prerequisites = [ordered]@{
        powershell = $null
        runtimes = @()
    }
    publish = [ordered]@{
        path = $resolvedPublishPath
        presentation_exe = $presentationExe
        worker_exe = $workerExe
        appsettings = $appSettingsPath
        presentation_process_id = $null
        worker_process_id = $null
        ui_connected_log_confirmed = $false
        supervisor_log_confirmed = $false
        worker_log_confirmed = $false
        launched_at_utc = $null
    }
    installer = [ordered]@{
        path = $resolvedSetupExePath
        product_name = "OnlyExo365"
        manufacturer = "OnlyExo365"
        product_version = $setupVersion.Split('+')[0].Trim()
        install_exit_code = $null
        uninstall_exit_code = $null
        install_root = $null
        required_files = @()
        installed_required_files = @()
        real_install_launch_confirmed = $false
        installed_presentation_process_id = $null
        installed_worker_process_id = $null
        uninstall_registry_confirmed = $false
        uninstall_cleanup_confirmed = $false
        residual_cleanup_confirmed = $false
        residual_marker_paths = @()
        local_data_backup_entries = @()
        local_data_restore_confirmed = $false
        skip_real_install = [bool]$SkipRealInstall
        quiet_uninstall_string = $null
    }
}

$installerSmokeRoot = $null
$productDataBackupRoot = $null
$productDataBackupEntries = @()

try {
    Write-Step "Validating runtime prerequisites"
    $report["prerequisites"] = Get-PrerequisiteReport -PresentationRuntimeConfig $presentationRuntimeConfig -WorkerRuntimeConfig $workerRuntimeConfig
    Write-Success "Runtime prerequisites confirmed for packaged execution."

    if (-not $SkipRealInstall) {
        $productDataBackup = Backup-ProductDataRoots -Paths @(
            $logDirectory,
            $secretDirectory,
            $exportDirectory
        )
        $productDataBackupRoot = $productDataBackup.BackupRoot
        $productDataBackupEntries = @($productDataBackup.Entries)
        $report["installer"]["local_data_backup_entries"] = @($productDataBackupEntries)
        if ($productDataBackupEntries.Count -gt 0) {
            Write-Success "Backed up $($productDataBackupEntries.Count) OnlyExo365 local data root(s) before installer validation."
        }
    }

    $publishLaunchReport = Invoke-ApplicationSmokeTest `
        -Label "publish" `
        -ExecutablePath $presentationExe `
        -WorkerExecutablePath $workerExe `
        -WorkingDirectory $resolvedPublishPath `
        -LogDirectory $logDirectory `
        -StartupTimeoutSeconds $StartupTimeoutSeconds `
        -StabilityWindowSeconds $StabilityWindowSeconds `
        -LogProbeTimeoutSeconds $LogProbeTimeoutSeconds
    $report["publish"] = [ordered]@{
        path = $resolvedPublishPath
        presentation_exe = $presentationExe
        worker_exe = $workerExe
        appsettings = $appSettingsPath
        presentation_process_id = $publishLaunchReport.presentation_process_id
        worker_process_id = $publishLaunchReport.worker_process_id
        ui_connected_log_confirmed = $publishLaunchReport.ui_connected_log_confirmed
        supervisor_log_confirmed = $publishLaunchReport.supervisor_log_confirmed
        worker_log_confirmed = $publishLaunchReport.worker_log_confirmed
        launched_at_utc = $publishLaunchReport.launched_at_utc
    }

    if (-not $SkipRealInstall) {
        $preExistingInstall = Get-InstalledProductEntry `
            -DisplayName $report["installer"]["product_name"] `
            -Publisher $report["installer"]["manufacturer"] `
            -InstallLocation $null
        if ($null -ne $preExistingInstall) {
            Stop-WithError "Existing installation detected for '$($report["installer"]["product_name"])'. Real installer smoke tests require a clean host to avoid altering an operator installation."
        }

        $installerSmokeRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("onlyexo365-setup-smoke-" + [guid]::NewGuid().ToString("N"))
        $realInstallRoot = Join-Path $installerSmokeRoot "installed\OnlyExo365"
        New-Item -Path $realInstallRoot -ItemType Directory -Force | Out-Null
        $report["installer"]["install_root"] = $realInstallRoot

        Write-Step "Installing setup EXE to temporary root"
        $installProcess = Start-Process -FilePath $resolvedSetupExePath -ArgumentList @(
                "/VERYSILENT",
                "/SUPPRESSMSGBOXES",
                "/NORESTART",
                "/DIR=$realInstallRoot"
            ) -Wait -PassThru

        $report["installer"]["install_exit_code"] = $installProcess.ExitCode
        if ($installProcess.ExitCode -ne 0) {
            Stop-WithError "Setup EXE installation failed (exit code: $($installProcess.ExitCode))."
        }

        $installedPayloadFiles = Get-RequiredPayloadFiles -RootPath $realInstallRoot
        Wait-Until -TimeoutSeconds $InstallerOperationTimeoutSeconds -FailureMessage "Installed payload did not materialize at $realInstallRoot" -Condition {
            foreach ($path in $installedPayloadFiles) {
                if (-not (Test-Path $path -PathType Leaf)) {
                    return $null
                }
            }

            return $true
        } | Out-Null

        $installedProduct = Wait-Until -TimeoutSeconds $InstallerOperationTimeoutSeconds -FailureMessage "Installed product entry did not appear in the uninstall registry." -Condition {
            Get-InstalledProductEntry -DisplayName $report["installer"]["product_name"] -Publisher $report["installer"]["manufacturer"] -InstallLocation $realInstallRoot
        }

        $report["installer"]["installed_required_files"] = $installedPayloadFiles
        $report["installer"]["quiet_uninstall_string"] = if (-not [string]::IsNullOrWhiteSpace($installedProduct.QuietUninstallString)) {
            $installedProduct.QuietUninstallString
        }
        else {
            $installedProduct.UninstallString
        }
        Write-Success "Installer EXE install produced the expected payload and uninstall registry entry."

        Write-Step "Reinstalling setup EXE over the existing installation"
        $reinstallProcess = Start-Process -FilePath $resolvedSetupExePath -ArgumentList @(
                "/VERYSILENT",
                "/SUPPRESSMSGBOXES",
                "/NORESTART",
                "/DIR=$realInstallRoot"
            ) -Wait -PassThru

        if ($reinstallProcess.ExitCode -ne 0) {
            Stop-WithError "Setup EXE reinstall failed (exit code: $($reinstallProcess.ExitCode))."
        }

        $installedProductEntriesAfterReinstall = @(Get-InstalledProductEntriesForLocation `
                -DisplayName $report["installer"]["product_name"] `
                -Publisher $report["installer"]["manufacturer"] `
                -InstallLocation $realInstallRoot)

        if ($installedProductEntriesAfterReinstall.Count -ne 1) {
            Stop-WithError "Setup EXE reinstall produced $($installedProductEntriesAfterReinstall.Count) uninstall registry entries for the installed product."
        }

        $installedProduct = $installedProductEntriesAfterReinstall | Select-Object -First 1
        $report["installer"]["quiet_uninstall_string"] = if (-not [string]::IsNullOrWhiteSpace($installedProduct.QuietUninstallString)) {
            $installedProduct.QuietUninstallString
        }
        else {
            $installedProduct.UninstallString
        }
        Write-Success "Setup reinstall preserved a single uninstall registry entry."

        $installedPresentationExe = Join-Path $realInstallRoot "OnlyExo365.Shell.exe"
        $installedWorkerExe = Join-Path $realInstallRoot "OnlyExo365.Worker.exe"
        $installedLaunchReport = Invoke-ApplicationSmokeTest `
            -Label "installed setup" `
            -ExecutablePath $installedPresentationExe `
            -WorkerExecutablePath $installedWorkerExe `
            -WorkingDirectory $realInstallRoot `
            -LogDirectory $logDirectory `
            -StartupTimeoutSeconds $StartupTimeoutSeconds `
            -StabilityWindowSeconds $StabilityWindowSeconds `
            -LogProbeTimeoutSeconds $LogProbeTimeoutSeconds

        $report["installer"]["real_install_launch_confirmed"] = $true
        $report["installer"]["installed_presentation_process_id"] = $installedLaunchReport.presentation_process_id
        $report["installer"]["installed_worker_process_id"] = $installedLaunchReport.worker_process_id

        $residualMarkerPaths = @(
            (New-ResidualMarkerFile -RootPath $realInstallRoot -RelativePath "residual\marker.txt"),
            (New-ResidualMarkerFile -RootPath $logDirectory -RelativePath "residual\marker.txt"),
            (New-ResidualMarkerFile -RootPath $secretDirectory -RelativePath "residual\marker.txt"),
            (New-ResidualMarkerFile -RootPath $exportDirectory -RelativePath "residual\marker.txt")
        )
        $report["installer"]["residual_marker_paths"] = $residualMarkerPaths

        Write-Step "Uninstalling setup EXE payload"
        $uninstallCommandLine = if (-not [string]::IsNullOrWhiteSpace($installedProduct.QuietUninstallString)) {
            $installedProduct.QuietUninstallString
        }
        else {
            $installedProduct.UninstallString
        }

        $uninstallInvocation = Get-UninstallInvocation -CommandLine $uninstallCommandLine
        $uninstallArguments = [string]$uninstallInvocation.arguments
        if ([string]::IsNullOrWhiteSpace($installedProduct.QuietUninstallString)) {
            $uninstallArguments = ($uninstallArguments + " /VERYSILENT /SUPPRESSMSGBOXES /NORESTART").Trim()
        }

        $uninstallProcess = Start-Process -FilePath $uninstallInvocation.file_path -ArgumentList $uninstallArguments -Wait -PassThru
        $report["installer"]["uninstall_exit_code"] = $uninstallProcess.ExitCode
        if ($uninstallProcess.ExitCode -ne 0) {
            Stop-WithError "Setup EXE uninstall failed (exit code: $($uninstallProcess.ExitCode))."
        }

        Wait-Until -TimeoutSeconds $InstallerOperationTimeoutSeconds -FailureMessage "Uninstall registry entry still present after setup uninstall." -Condition {
            $entry = Get-InstalledProductEntry -DisplayName $report["installer"]["product_name"] -Publisher $report["installer"]["manufacturer"] -InstallLocation $realInstallRoot
            if ($null -eq $entry) {
                return $true
            }

            return $null
        } | Out-Null

        Wait-Until -TimeoutSeconds $InstallerOperationTimeoutSeconds -FailureMessage "Installed payload still present after setup uninstall." -Condition {
            foreach ($path in $installedPayloadFiles) {
                if (Test-Path $path -PathType Leaf) {
                    return $null
                }
            }

            return $true
        } | Out-Null

        Wait-Until -TimeoutSeconds $InstallerOperationTimeoutSeconds -FailureMessage "Residual OnlyExo365 files or directories are still present after setup uninstall." -Condition {
            foreach ($path in $residualMarkerPaths) {
                if (Test-Path $path -PathType Leaf) {
                    return $null
                }
            }

            foreach ($root in @($realInstallRoot, $logDirectory, $secretDirectory, $exportDirectory)) {
                if (Test-Path $root) {
                    return $null
                }
            }

            return $true
        } | Out-Null

        $report["installer"]["required_files"] = Get-RequiredPayloadFiles -RootPath $realInstallRoot
        $report["installer"]["uninstall_registry_confirmed"] = $true
        $report["installer"]["uninstall_cleanup_confirmed"] = $true
        $report["installer"]["residual_cleanup_confirmed"] = $true
        Write-Success "Setup uninstall removed the registered installation and residual files."
    }
    else {
        Write-Info "Skipping real setup EXE install validation by request."
    }

    if ($productDataBackupEntries.Count -gt 0) {
        $restoredCount = Restore-ProductDataRoots -Entries $productDataBackupEntries
        if ($restoredCount -ne $productDataBackupEntries.Count) {
            Stop-WithError "Unable to restore all pre-existing OnlyExo365 local data after installer validation."
        }

        $productDataBackupEntries = @()
        $report["installer"]["local_data_restore_confirmed"] = $true
        Write-Success "Restored pre-existing OnlyExo365 local data after installer validation."
    }

    $report | ConvertTo-Json -Depth 10 | Out-File -FilePath $resolvedReportPath -Encoding utf8
    Write-Info "Smoke report: $resolvedReportPath"
}
finally {
    if ($null -ne $installerSmokeRoot -and (Test-Path $installerSmokeRoot)) {
        Remove-Item -Path $installerSmokeRoot -Recurse -Force -ErrorAction SilentlyContinue
    }

    if ($productDataBackupEntries.Count -gt 0) {
        try {
            $restoredCount = Restore-ProductDataRoots -Entries $productDataBackupEntries
            if ($restoredCount -eq $productDataBackupEntries.Count) {
                Write-Warn "Recovered pre-existing OnlyExo365 local data during final cleanup after an interrupted installer smoke test run."
                $productDataBackupEntries = @()
            }
            else {
                Write-Warn "Smoke test cleanup could not restore all pre-existing OnlyExo365 local data automatically."
            }
        }
        catch {
            Write-Warn "Smoke test cleanup failed while restoring pre-existing OnlyExo365 local data: $($_.Exception.Message)"
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($productDataBackupRoot) -and (Test-Path $productDataBackupRoot)) {
        Remove-Item -Path $productDataBackupRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " SMOKE TESTS COMPLETED" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

