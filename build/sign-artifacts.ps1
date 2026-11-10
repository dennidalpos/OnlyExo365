#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string[]]$Path,

    [Parameter(Mandatory = $true)]
    [string]$CertificatePath,

    [string]$CertificatePassword,

    [ValidateSet("SHA256", "SHA384", "SHA512")]
    [string]$DigestAlgorithm = "SHA256",

    [string]$TimestampUrl = "http://timestamp.digicert.com",

    [switch]$AllowUntrustedSigner
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

. (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "signing-helpers.ps1")
. (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "resolve-signtool.ps1")

function Get-SignToolPath {
    try {
        return Resolve-SignToolPath
    }
    catch {
        Stop-WithError $_.Exception.Message
    }
}

function Test-SignedFile {
    param(
        [string]$TargetPath,
        [string]$ExpectedThumbprint,
        [bool]$AllowUntrustedSigner
    )

    $signature = Get-AuthenticodeSignature -FilePath $TargetPath
    $actualThumbprint = $signature.SignerCertificate?.Thumbprint
    $thumbprintMatches = -not [string]::IsNullOrWhiteSpace($actualThumbprint) -and
        $actualThumbprint.Equals($ExpectedThumbprint, [System.StringComparison]::OrdinalIgnoreCase)

    if ($signature.Status -eq [System.Management.Automation.SignatureStatus]::Valid) {
        return $thumbprintMatches
    }

    return $AllowUntrustedSigner -and
        ($signature.Status -eq [System.Management.Automation.SignatureStatus]::NotTrusted -or
         $signature.Status -eq [System.Management.Automation.SignatureStatus]::UnknownError) -and
        $thumbprintMatches
}

if (-not (Test-Path $CertificatePath -PathType Leaf)) {
    Stop-WithError "Certificate file not found: $CertificatePath"
}

$TimestampUrl = Resolve-TimestampUrl -Url $TimestampUrl
$signingCertificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
    $CertificatePath,
    $CertificatePassword,
    [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::DefaultKeySet)
$expectedThumbprint = $signingCertificate.Thumbprint
$signTool = Get-SignToolPath
$targets = Get-SignableTargets -InputPaths $Path -UnsupportedFileMessage "Unsupported file type for signing: {0}"

if ($targets.Count -eq 0) {
    Stop-WithError "No .exe files found to sign."
}

Write-Step "Signing artifacts"
Write-Info "signtool : $signTool"
Write-Info "Targets  : $($targets.Count)"
Write-Info "Cert     : $CertificatePath"
Write-Info "Digest   : $DigestAlgorithm"
Write-Info "Timestamp: $TimestampUrl"

foreach ($target in $targets) {
    $arguments = @(
        "sign",
        "/fd", $DigestAlgorithm,
        "/td", $DigestAlgorithm,
        "/tr", $TimestampUrl,
        "/f", $CertificatePath
    )

    if (-not [string]::IsNullOrWhiteSpace($CertificatePassword)) {
        $arguments += "/p", $CertificatePassword
    }

    $arguments += "/v", $target

    Write-Info "Signing $target"
    & $signTool @arguments
    if ($LASTEXITCODE -ne 0) {
        Stop-WithError "signtool failed while signing $target (exit code: $LASTEXITCODE)" $LASTEXITCODE
    }

    if (-not (Test-SignedFile -TargetPath $target -ExpectedThumbprint $expectedThumbprint -AllowUntrustedSigner ([bool]$AllowUntrustedSigner))) {
        Stop-WithError "Signing completed without a valid Authenticode signature on $target"
    }

    Write-Success "Signed: $target"
}

$signingCertificate.Dispose()

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " SIGNING COMPLETED" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
