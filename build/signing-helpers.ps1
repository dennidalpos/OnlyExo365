#Requires -Version 7.0

. (Join-Path (Split-Path -Parent $PSScriptRoot) "scripts\internal\common.ps1")

function Resolve-TimestampUrl {
    param([string]$Url)

    if ([string]::IsNullOrWhiteSpace($Url)) {
        Stop-WithError "TimestampUrl must be provided."
    }

    $uri = $null
    if (-not [System.Uri]::TryCreate($Url.Trim(), [System.UriKind]::Absolute, [ref]$uri)) {
        Stop-WithError "TimestampUrl must be an absolute URL: $Url"
    }

    if ($uri.Scheme -ne [System.Uri]::UriSchemeHttps -and $uri.Scheme -ne [System.Uri]::UriSchemeHttp) {
        Stop-WithError "TimestampUrl must use HTTP or HTTPS: $Url"
    }

    $trimmedUrl = $Url.Trim()
    $normalizedUrl = $uri.AbsoluteUri

    if ($uri.AbsolutePath -eq "/" -and -not $trimmedUrl.EndsWith("/")) {
        return $normalizedUrl.TrimEnd('/')
    }

    return $normalizedUrl
}

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

function Get-SignableTargets {
    param(
        [string[]]$InputPaths,
        [string]$UnsupportedFileMessage
    )

    $targets = New-Object System.Collections.Generic.List[string]

    foreach ($rawItem in $InputPaths) {
        foreach ($item in ($rawItem -split ',')) {
            if ([string]::IsNullOrWhiteSpace($item)) {
                continue
            }

            $trimmedItem = $item.Trim()
            $resolved = Resolve-Path -Path $trimmedItem -ErrorAction Stop
            foreach ($entry in $resolved) {
                if (Test-Path $entry.Path -PathType Container) {
                    Get-ChildItem -Path $entry.Path -Recurse -File -Include *.exe |
                        Sort-Object FullName |
                        ForEach-Object { [void]$targets.Add($_.FullName) }
                }
                else {
                    $extension = [System.IO.Path]::GetExtension($entry.Path)
                    if ($extension -ne ".exe") {
                        Stop-WithError ($UnsupportedFileMessage -f $entry.Path)
                    }

                    [void]$targets.Add($entry.Path)
                }
            }
        }
    }

    return @($targets | Sort-Object -Unique)
}
