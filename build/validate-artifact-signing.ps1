#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string[]]$Path,

    [string]$SignerSubject = "CN=ExchangeAdmin CI Validation",
    [string]$TimestampUrl = "http://timestamp.digicert.com"
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

. (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "signing-helpers.ps1")

function Remove-CertificateFromStore {
    param(
        [string]$StorePath,
        [string]$Thumbprint,
        [switch]$DeleteKey
    )

    if (-not (Test-Path $StorePath)) {
        return
    }

    Get-ChildItem -Path $StorePath -ErrorAction SilentlyContinue |
        Where-Object { $_.Thumbprint -eq $Thumbprint } |
        ForEach-Object {
            try {
                $removeArguments = @{
                    Path = $_.PSPath
                    Force = $true
                    ErrorAction = 'Stop'
                }

                if ($DeleteKey) {
                    $removeArguments.DeleteKey = $true
                }

                Remove-Item @removeArguments
            }
            catch {
                Write-Verbose "Skipped certificate cleanup for $StorePath ($Thumbprint): $($_.Exception.Message)"
            }
        }
}

function Remove-CertificatesBySubject {
    param(
        [string]$StorePath,
        [string]$Subject,
        [switch]$DeleteKey
    )

    if ([string]::IsNullOrWhiteSpace($Subject) -or -not (Test-Path $StorePath)) {
        return
    }

    Get-ChildItem -Path $StorePath -ErrorAction SilentlyContinue |
        Where-Object { $_.Subject -eq $Subject } |
        ForEach-Object {
            try {
                $removeArguments = @{
                    Path = $_.PSPath
                    Force = $true
                    ErrorAction = 'Stop'
                }

                if ($DeleteKey) {
                    $removeArguments.DeleteKey = $true
                }

                Remove-Item @removeArguments
            }
            catch {
                Write-Verbose "Skipped certificate cleanup for $StorePath ($($_.Thumbprint)): $($_.Exception.Message)"
            }
        }
}

$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$repositoryRoot = Split-Path -Parent $scriptDirectory
$signScript = Join-Path $scriptDirectory "sign-artifacts.ps1"
$verifyScript = Join-Path $scriptDirectory "verify-signatures.ps1"

if (-not (Test-Path $signScript -PathType Leaf)) {
    Stop-WithError "Signing script not found: $signScript"
}

if (-not (Test-Path $verifyScript -PathType Leaf)) {
    Stop-WithError "Signature verification script not found: $verifyScript"
}

$resolvedPaths = @()
foreach ($entry in $Path) {
    foreach ($item in ($entry -split ',')) {
        if ([string]::IsNullOrWhiteSpace($item)) {
            continue
        }

        $resolvedPath = Resolve-RepoPath -BaseDirectory $repositoryRoot -PathValue $item.Trim()
        if (-not (Test-Path $resolvedPath)) {
            Stop-WithError "Input path not found: $resolvedPath"
        }

        $resolvedPaths += $resolvedPath
    }
}

$temporaryRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("exchangeadmin-signing-validation-" + [guid]::NewGuid().ToString("N"))
$certificatePath = Join-Path $temporaryRoot "validation-signing.pfx"
$clonedTargets = New-Object System.Collections.Generic.List[string]
$certificateThumbprint = $null
$TimestampUrl = Resolve-TimestampUrl -Url $TimestampUrl

try {
    New-Item -Path $temporaryRoot -ItemType Directory -Force | Out-Null

    # Remove stale disposable validation certificates from previous interrupted runs.
    Remove-CertificatesBySubject -StorePath "Cert:\CurrentUser\My" -Subject $SignerSubject -DeleteKey
    Write-Step "Preparing disposable artifact copies for signature validation"

    for ($index = 0; $index -lt $resolvedPaths.Count; $index++) {
        $sourcePath = $resolvedPaths[$index]
        $sourceItem = Get-Item -Path $sourcePath
        $destinationPath = Join-Path $temporaryRoot ("input-{0}-{1}" -f $index, $sourceItem.Name)

        if ($sourceItem.PSIsContainer) {
            Copy-Item -Path $sourceItem.FullName -Destination $destinationPath -Recurse -Force
        }
        else {
            Copy-Item -Path $sourceItem.FullName -Destination $destinationPath -Force
        }

        [void]$clonedTargets.Add($destinationPath)
        Write-Info "Copied $($sourceItem.FullName) -> $destinationPath"
    }

    Write-Step "Creating disposable code-signing certificate"
    $certificatePasswordPlain = [guid]::NewGuid().ToString("N") + "!"
    $certificatePassword = ConvertTo-SecureString -String $certificatePasswordPlain -AsPlainText -Force

    $certificate = New-SelfSignedCertificate `
        -Subject $SignerSubject `
        -Type Custom `
        -KeyAlgorithm RSA `
        -KeyLength 2048 `
        -HashAlgorithm SHA256 `
        -KeyExportPolicy Exportable `
        -CertStoreLocation "Cert:\CurrentUser\My" `
        -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.3") `
        -NotAfter (Get-Date).AddDays(30)

    if ($null -eq $certificate) {
        Stop-WithError "Unable to create disposable signing certificate."
    }

    $certificateThumbprint = $certificate.Thumbprint

    Export-PfxCertificate -Cert $certificate -FilePath $certificatePath -Password $certificatePassword | Out-Null

    Write-Info "Disposable certificate thumbprint: $certificateThumbprint"

    Write-Step "Signing disposable artifact copies"
    & $signScript `
        -Path $clonedTargets.ToArray() `
        -CertificatePath $certificatePath `
        -CertificatePassword $certificatePasswordPlain `
        -TimestampUrl $TimestampUrl `
        -AllowUntrustedSigner

    if ($LASTEXITCODE -ne 0) {
        Stop-WithError "Artifact signing validation failed during sign-artifacts execution." $LASTEXITCODE
    }

    Write-Step "Verifying disposable artifact signatures"
    & $verifyScript `
        -Path $clonedTargets.ToArray() `
        -ExpectedThumbprint $certificateThumbprint `
        -AllowUntrustedSigner

    if ($LASTEXITCODE -ne 0) {
        Stop-WithError "Artifact signing validation failed during verify-signatures execution." $LASTEXITCODE
    }

    Write-Success "Disposable artifact signing path validated successfully."
}
finally {
    if (-not [string]::IsNullOrWhiteSpace($certificateThumbprint)) {
        Remove-CertificateFromStore -StorePath "Cert:\CurrentUser\My" -Thumbprint $certificateThumbprint -DeleteKey
    }

    Remove-CertificatesBySubject -StorePath "Cert:\CurrentUser\My" -Subject $SignerSubject -DeleteKey

    if (Test-Path $temporaryRoot) {
        Remove-Item -Path $temporaryRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
