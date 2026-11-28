#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string[]]$Path,

    [string]$ExpectedThumbprint,

    [switch]$AllowUntrustedSigner
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

. (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "signing-helpers.ps1")

$targets = Get-SignableTargets -InputPaths $Path -UnsupportedFileMessage "Unsupported file type for signature verification: {0}"
if ($targets.Count -eq 0) {
    Stop-WithError "No .exe files found to verify."
}

$normalizedThumbprint = if ([string]::IsNullOrWhiteSpace($ExpectedThumbprint)) {
    $null
}
else {
    $ExpectedThumbprint.Trim().ToUpperInvariant()
}

Write-Step "Verifying Authenticode signatures"
Write-Info "Targets: $($targets.Count)"

foreach ($target in $targets) {
    $signature = Get-AuthenticodeSignature -FilePath $target
    $isAcceptedStatus =
        $signature.Status -eq [System.Management.Automation.SignatureStatus]::Valid -or
        ($AllowUntrustedSigner -and
            ($signature.Status -eq [System.Management.Automation.SignatureStatus]::NotTrusted -or
             $signature.Status -eq [System.Management.Automation.SignatureStatus]::UnknownError))

    if (-not $isAcceptedStatus) {
        Stop-WithError "Invalid or missing Authenticode signature on $target. Status: $($signature.Status)"
    }

    if ($null -ne $normalizedThumbprint) {
        $actualThumbprint = $signature.SignerCertificate?.Thumbprint
        if ([string]::IsNullOrWhiteSpace($actualThumbprint)) {
            Stop-WithError "Signed file $target does not expose a signer certificate thumbprint."
        }

        if ($actualThumbprint.ToUpperInvariant() -ne $normalizedThumbprint) {
            Stop-WithError "Unexpected signer thumbprint on $target. Expected: $normalizedThumbprint, actual: $actualThumbprint"
        }
    }

    Write-Success "Valid signature: $target"
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " SIGNATURE VERIFICATION COMPLETED" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
