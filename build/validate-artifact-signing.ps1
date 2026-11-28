#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string[]]$Path,

    [string]$SignerSubject = "CN=OnlyExo365 CI Validation",
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

    Get-ChildItem -LiteralPath $StorePath -ErrorAction Stop |
        Where-Object { $_.Thumbprint -eq $Thumbprint } |
        ForEach-Object {
            $removeArguments = @{
                LiteralPath = $_.PSPath
                Force = $true
                ErrorAction = 'Stop'
            }

            if ($DeleteKey) {
                $removeArguments.DeleteKey = $true
            }

            Remove-Item @removeArguments
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

        $resolvedPath = Resolve-RepositoryPath -RepositoryRoot $repositoryRoot -PathValue $item.Trim()
        if (-not (Test-Path $resolvedPath)) {
            Stop-WithError "Input path not found: $resolvedPath"
        }

        $resolvedPaths += $resolvedPath
    }
}

$temporaryRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("onlyexo365-signing-validation-" + [guid]::NewGuid().ToString("N"))
$certificatePath = Join-Path $temporaryRoot "validation-signing.pfx"
$clonedTargets = New-Object System.Collections.Generic.List[string]
$certificateThumbprint = $null
$certificatePassword = $null
$TimestampUrl = Resolve-TimestampUrl -Url $TimestampUrl

try {
    New-Item -Path $temporaryRoot -ItemType Directory -Force | Out-Null

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
    $certificatePassword = [System.Security.SecureString]::new()
    foreach ($character in $certificatePasswordPlain.ToCharArray()) {
        $certificatePassword.AppendChar($character)
    }

    $certificatePassword.MakeReadOnly()

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
    if ($null -ne $certificatePassword) {
        $certificatePassword.Dispose()
    }

    if (-not [string]::IsNullOrWhiteSpace($certificateThumbprint)) {
        Remove-CertificateFromStore -StorePath "Cert:\CurrentUser\My" -Thumbprint $certificateThumbprint -DeleteKey
    }

    if (Test-Path -LiteralPath $temporaryRoot) {
        Remove-Item -LiteralPath $temporaryRoot -Recurse -Force -ErrorAction Stop
    }
}

