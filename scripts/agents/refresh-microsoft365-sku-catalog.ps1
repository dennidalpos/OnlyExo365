#Requires -Version 7.0

[CmdletBinding()]
param(
    [ValidateNotNullOrEmpty()]
    [string]$DocumentationUrl = "https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference",

    [string]$CsvDownloadUrl,

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = "src/OnlyExo365.Worker/Data/Microsoft365SkuCatalog.json",

    [ValidateNotNullOrEmpty()]
    [string]$WorkingDirectory = "artifacts/tmp/licensing-catalog"
)

. (Join-Path $PSScriptRoot "../internal/common.ps1")

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$repositoryRoot = Get-RepositoryRoot -ScriptRoot $PSScriptRoot
$resolvedOutputPath = Resolve-RepositoryPath -RepositoryRoot $repositoryRoot -PathValue $OutputPath
$resolvedWorkingDirectory = Resolve-RepositoryPath -RepositoryRoot $repositoryRoot -PathValue $WorkingDirectory

function Get-LicensingCsvUrl {
    param([string]$SourceDocumentationUrl)

    Write-Step "Resolving Microsoft licensing CSV URL"
    Write-Info "Documentation: $SourceDocumentationUrl"

    $response = Invoke-WebRequest -Uri $SourceDocumentationUrl
    $downloadLink = $response.Links |
        Where-Object {
            $_.PSObject.Properties.Match('href').Count -gt 0 -and
            $_.href -match '^https://download\.microsoft\.com/.+licensing\.csv(?:\?|$)'
        } |
        Select-Object -First 1 -ExpandProperty href

    if (-not [string]::IsNullOrWhiteSpace($downloadLink)) {
        return $downloadLink
    }

    $contentMatch = [regex]::Match(
        $response.Content,
        'https://download\.microsoft\.com/download/[^"\s<>]+licensing\.csv',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

    if ($contentMatch.Success) {
        return $contentMatch.Value
    }

    throw "Unable to resolve the Microsoft licensing CSV link from $SourceDocumentationUrl"
}

function Normalize-String {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return [string]::Empty
    }

    return $Value.Trim()
}

function Convert-CsvRowsToCatalogEntries {
    param([object[]]$Rows)

    if ($Rows.Count -eq 0) { throw 'Licensing CSV contains no entries.' }
    $columns = @('String_Id', 'GUID', 'Product_Display_Name', 'Service_Plan_Name', 'Service_Plan_Id', 'Service_Plans_Included_Friendly_Names')
    foreach ($row in $Rows) {
        foreach ($column in $columns) {
            if ($null -eq $row.PSObject.Properties[$column]) { throw "Licensing CSV missing column: $column" }
        }
        $identifier = [guid]::Empty
        if (-not [guid]::TryParse([string]$row.GUID, [ref]$identifier) -or $identifier -eq [guid]::Empty -or
            [string]::IsNullOrWhiteSpace($row.String_Id) -or [string]::IsNullOrWhiteSpace($row.Product_Display_Name)) {
            throw 'Licensing CSV contains an invalid SKU identity or product name.'
        }
        if (-not [string]::IsNullOrWhiteSpace($row.Service_Plan_Id) -and
            (-not [guid]::TryParse([string]$row.Service_Plan_Id, [ref]$identifier) -or $identifier -eq [guid]::Empty)) {
            throw 'Licensing CSV contains an invalid service plan identity.'
        }
    }

    $groupedRows = $Rows |
        Group-Object {
            "{0}|{1}" -f (Normalize-String -Value $_.String_Id), (Normalize-String -Value $_.GUID)
        }

    $entries = foreach ($group in $groupedRows) {
        $firstRow = $group.Group | Select-Object -First 1
        $servicePlans = foreach ($row in ($group.Group | Sort-Object Service_Plan_Name, Service_Plan_Id)) {
            $servicePlanName = Normalize-String -Value $row.Service_Plan_Name
            $servicePlanId = Normalize-String -Value $row.Service_Plan_Id
            $friendlyName = Normalize-String -Value $row.Service_Plans_Included_Friendly_Names

            if ([string]::IsNullOrWhiteSpace($servicePlanName) -and [string]::IsNullOrWhiteSpace($servicePlanId) -and [string]::IsNullOrWhiteSpace($friendlyName)) {
                continue
            }

            [ordered]@{
                servicePlanName = $servicePlanName
                servicePlanId = $servicePlanId
                friendlyName = $friendlyName
            }
        }

        [ordered]@{
            skuId = Normalize-String -Value $firstRow.GUID
            skuPartNumber = Normalize-String -Value $firstRow.String_Id
            productName = Normalize-String -Value $firstRow.Product_Display_Name
            servicePlans = @($servicePlans)
        }
    }

    return @($entries | Sort-Object skuPartNumber, skuId)
}

New-Item -Path $resolvedWorkingDirectory -ItemType Directory -Force | Out-Null
$downloadedCsvPath = Join-Path $resolvedWorkingDirectory "microsoft365-licensing.csv"
$effectiveCsvDownloadUrl = if ([string]::IsNullOrWhiteSpace($CsvDownloadUrl)) {
    Get-LicensingCsvUrl -SourceDocumentationUrl $DocumentationUrl
}
else {
    $CsvDownloadUrl
}

Write-Step "Downloading Microsoft licensing catalog"
Write-Info "CSV: $effectiveCsvDownloadUrl"
Invoke-WebRequest -Uri $effectiveCsvDownloadUrl -OutFile $downloadedCsvPath

$rows = @(Import-Csv -Path $downloadedCsvPath)
if ($rows.Count -eq 0) {
    throw "The downloaded Microsoft licensing CSV is empty: $downloadedCsvPath"
}

$catalogDocument = [ordered]@{
    generatedOn = (Get-Date).ToString("yyyy-MM-dd")
    source = $DocumentationUrl
    csvDownload = $effectiveCsvDownloadUrl
    entries = @(Convert-CsvRowsToCatalogEntries -Rows $rows)
}

New-Item -Path (Split-Path -Parent $resolvedOutputPath) -ItemType Directory -Force | Out-Null
$temporaryOutput = $resolvedOutputPath + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
try {
    $catalogDocument | ConvertTo-Json -Depth 6 -Compress | Set-Content -LiteralPath $temporaryOutput -Encoding utf8NoBOM
    [IO.File]::Move($temporaryOutput, $resolvedOutputPath, $true)
}
finally {
    if (Test-Path -LiteralPath $temporaryOutput) { Remove-Item -LiteralPath $temporaryOutput -ErrorAction Stop }
}

Write-Success "Microsoft 365 SKU catalog refreshed: $resolvedOutputPath"
Write-Info "Entries: $($catalogDocument.entries.Count)"

