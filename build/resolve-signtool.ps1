#Requires -Version 7.0

[CmdletBinding()]
param(
    [switch]$DirectoryOnly
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
. (Join-Path (Split-Path -Parent $PSScriptRoot) "scripts\internal\common.ps1")

function Get-SignToolSearchDirectories {
    $kitRoots = @()

    $windowsKitBin = "C:\Program Files (x86)\Windows Kits\10\bin"
    if (Test-Path $windowsKitBin) {
        $kitRoots += Get-ChildItem -Path $windowsKitBin -Directory -ErrorAction SilentlyContinue |
            Sort-Object Name -Descending |
            ForEach-Object { Join-Path $_.FullName "x64" }
    }

    return @(
        $env:WINDOWS_SIGNTOOL_DIR,
        $env:WINDOWS_KITS_BIN,
        "C:\Program Files (x86)\Windows Kits\10\App Certification Kit"
    ) + $kitRoots
}

function Resolve-SignToolPath {
    $candidateDirectories = Get-SignToolSearchDirectories
    $toolPath = Get-ExecutablePath -CommandName "signtool.exe" -CandidateDirectories $candidateDirectories

    if ($null -ne $toolPath) {
        return $toolPath
    }

    $searched = ($candidateDirectories |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique) -join ", "

    throw "signtool.exe not found. Install Windows SDK / App Certification Kit or expose it via WINDOWS_SIGNTOOL_DIR. Paths checked: $searched"
}

if ($MyInvocation.InvocationName -ne '.') {
    $resolvedPath = Resolve-SignToolPath
    if ($DirectoryOnly) {
        Split-Path -Parent $resolvedPath
    }
    else {
        $resolvedPath
    }
}
