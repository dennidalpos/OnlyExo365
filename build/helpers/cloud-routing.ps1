function Resolve-CloudEnvironment {
    param([string]$ExchangeEnvironmentName)
    if ([string]::IsNullOrWhiteSpace($ExchangeEnvironmentName)) { $ExchangeEnvironmentName = 'O365Default' }
    $routes = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot '../../src/OnlyExo365.Contracts/Data/CloudEnvironments.json') | ConvertFrom-Json -AsHashtable
    if (-not $routes.ContainsKey($ExchangeEnvironmentName)) { throw "Unsupported Microsoft 365 cloud: '$ExchangeEnvironmentName'." }
    return $routes[$ExchangeEnvironmentName]
}

function Get-ComplianceRoutingArguments {
    param([hashtable]$Configuration)
    $route = Resolve-CloudEnvironment $Configuration.ExchangeEnvironmentName
    $tenant = if ([string]::IsNullOrWhiteSpace($Configuration.DelegatedOrganization)) { 'organizations' } else { [uri]::EscapeDataString($Configuration.DelegatedOrganization) }
    return @{ ConnectionUri = $route.ComplianceUri; AzureADAuthorizationEndpointUri = "$($route.Authority)/$tenant" }
}
