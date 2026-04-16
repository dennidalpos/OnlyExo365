#Requires -Version 7.0

[CmdletBinding()]
param(
    [switch]$Install,

    [ValidateSet("Auto", "Winget", "Chocolatey")]
    [string]$PackageManager = "Auto"
)

. (Join-Path $PSScriptRoot "helpers/common.ps1")

$repositoryRoot = Get-RepositoryRoot -ScriptRoot $PSScriptRoot

function Test-InnoSetupAvailable {
    param([string]$RepositoryRoot)

    try {
        $compilerPath = Get-InnoSetupCompilerPath -RepositoryRoot $RepositoryRoot
        $versionLine = (& $compilerPath "/?" 2>&1 | Select-Object -First 1).ToString().Trim()
        $exitCode = if (Get-Variable -Name LASTEXITCODE -Scope Global -ErrorAction SilentlyContinue) {
            $global:LASTEXITCODE
        }
        else {
            0
        }

        if ($exitCode -ne 0 -or $versionLine -notmatch "^Inno Setup 6\b") {
            throw "ISCC.exe was found, but it is not Inno Setup 6. Found: $versionLine"
        }

        return [pscustomobject]@{
            Found = $true
            CompilerPath = $compilerPath
            BinPath = Split-Path -Parent $compilerPath
            VersionLine = $versionLine
        }
    }
    catch {
        return [pscustomobject]@{
            Found = $false
            CompilerPath = $null
            BinPath = $null
            VersionLine = $null
            Error = $_.Exception.Message
        }
    }
}

function Invoke-PackageInstall {
    param(
        [string]$CommandName,
        [string[]]$Arguments,
        [string]$DisplayName
    )

    $command = Get-Command $CommandName -ErrorAction SilentlyContinue
    if ($null -eq $command) {
        return $false
    }

    Write-Step "Installing Inno Setup 6 with $DisplayName"
    Write-Info "Command: $CommandName $($Arguments -join ' ')"

    & $command.Source @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$DisplayName failed to install Inno Setup 6 (exit code: $LASTEXITCODE)."
    }

    return $true
}

Write-Step "Preparing Inno Setup 6 for local packaging"
Assert-WindowsPlatform

$existingInnoSetup = Test-InnoSetupAvailable -RepositoryRoot $repositoryRoot
if ($existingInnoSetup.Found) {
    Write-Info "Inno Setup bin: $($existingInnoSetup.BinPath)"
    Write-Info "ISCC: $($existingInnoSetup.CompilerPath)"
    Write-Info "Version: $($existingInnoSetup.VersionLine)"
    Write-Success "Inno Setup 6 is available"
    return
}

if (-not $Install) {
    Write-Warn $existingInnoSetup.Error
    throw "Inno Setup 6 is not available. Install it manually, set INNOSETUP_BIN/INNOSETUP_HOME, or rerun with -Install."
}

$installAttempted = $false
$installErrors = @()

if ($PackageManager -in @("Auto", "Winget")) {
    try {
        $installAttempted = Invoke-PackageInstall `
            -CommandName "winget" `
            -DisplayName "winget" `
            -Arguments @(
                "install",
                "--id", "JRSoftware.InnoSetup",
                "--exact",
                "--source", "winget",
                "--accept-package-agreements",
                "--accept-source-agreements"
            )
    }
    catch {
        if ($PackageManager -eq "Winget") {
            throw
        }

        $installErrors += $_.Exception.Message
        Write-Warn $_.Exception.Message
    }
}

if (-not $installAttempted -and $PackageManager -in @("Auto", "Chocolatey")) {
    try {
        $installAttempted = Invoke-PackageInstall `
            -CommandName "choco" `
            -DisplayName "Chocolatey" `
            -Arguments @(
                "install",
                "innosetup",
                "--version=6.7.1",
                "-y",
                "--no-progress"
            )
    }
    catch {
        if ($PackageManager -eq "Chocolatey") {
            throw
        }

        $installErrors += $_.Exception.Message
        Write-Warn $_.Exception.Message
    }
}

if (-not $installAttempted) {
    if ($installErrors.Count -gt 0) {
        throw "Inno Setup 6 installation failed. Attempts: $($installErrors -join ' | ')"
    }

    throw "No supported package manager was found. Install Inno Setup 6 manually or configure INNOSETUP_BIN/INNOSETUP_HOME."
}

$installedInnoSetup = Test-InnoSetupAvailable -RepositoryRoot $repositoryRoot
if (-not $installedInnoSetup.Found) {
    throw "Inno Setup installation completed, but ISCC.exe was not found. Configure INNOSETUP_BIN/INNOSETUP_HOME if it was installed to a custom path."
}

Write-Info "Inno Setup bin: $($installedInnoSetup.BinPath)"
Write-Info "ISCC: $($installedInnoSetup.CompilerPath)"
Write-Info "Version: $($installedInnoSetup.VersionLine)"
Write-Success "Inno Setup 6 is ready for local packaging"
