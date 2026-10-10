#Requires -Version 7.6

[CmdletBinding()]
param(
    [string]$ConfigurationPath,
    [string]$ReportPath = "artifacts/tenant-validation/tenant-validation-report.json",
    [switch]$SkipGraph = $false,
    [switch]$InstallApprovedModules = $true
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$repositoryRoot = Split-Path -Parent $scriptDirectory

. (Join-Path $repositoryRoot "scripts\internal\common.ps1")

function Read-OptionalValue {
    param($Value)

    if ($null -eq $Value) {
        return $null
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    return $text.Trim()
}

function Read-OptionalBoolean {
    param($Value)

    if ($null -eq $Value) {
        return $null
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    return $text.Trim().ToLowerInvariant() -in @("1", "true", "yes", "on")
}

function Read-StringList {
    param($Value)

    if ($null -eq $Value) {
        return @()
    }

    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        return @($Value | ForEach-Object { Read-OptionalValue $_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }

    return @([string]$Value -split ';' | ForEach-Object { Read-OptionalValue $_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

function Get-EnvValue {
    param([string]$Name)

    return Read-OptionalValue ([Environment]::GetEnvironmentVariable($Name))
}

function Test-NonUiPowerShellHost {
    if ($env:GITHUB_ACTIONS -eq "true" -or $env:CI -eq "true") {
        return $true
    }

    if (-not [Environment]::UserInteractive) {
        return $true
    }

    try {
        $process = Get-Process -Id $PID -ErrorAction Stop
        if ($process.MainWindowHandle -eq 0) {
            return $true
        }
    }
    catch {
        return $true
    }

    return $false
}

function Resolve-EffectiveAuthenticationMode {
    param([hashtable]$Configuration)

    $requestedMode = $Configuration.AuthenticationMode.Trim()
    $effectiveMode = $requestedMode
    $fallbackReason = $null

    if ($requestedMode.Equals("Interactive", [System.StringComparison]::OrdinalIgnoreCase) -and (Test-NonUiPowerShellHost)) {
        $effectiveMode = "DeviceCode"
        $fallbackReason = "Interactive authentication requested from a non-UI PowerShell host. Falling back to DeviceCode to avoid the WAM failure 'A window handle must be configured'."
        Write-Info $fallbackReason
    }

    return [ordered]@{
        requested_mode = $requestedMode
        effective_mode = $effectiveMode
        used_device_code_fallback = $effectiveMode -ne $requestedMode
        fallback_reason = $fallbackReason
    }
}

function Merge-StringValue {
    param(
        [hashtable]$Target,
        [string]$Key,
        [object]$Value
    )

    $normalized = Read-OptionalValue $Value
    if (-not [string]::IsNullOrWhiteSpace($normalized)) {
        $Target[$Key] = $normalized
    }
}

function Merge-StringListValue {
    param(
        [hashtable]$Target,
        [string]$Key,
        [object]$Value
    )

    $items = @(Read-StringList $Value)
    if ($items.Count -gt 0) {
        $Target[$Key] = $items | Select-Object -Unique
    }
}

function Load-ExchangeConfiguration {
    param([string]$Path)

    $configuration = @{
        AuthenticationMode = "Interactive"
        ExchangeEnvironmentName = "O365Default"
        ExchangeOrganization = $null
        DelegatedOrganization = $null
        UserPrincipalNameHint = $null
        ApplicationId = $null
        CertificateThumbprint = $null
        CertificateSubjectName = $null
        ManagedIdentityAccountId = $null
        GraphTenantId = $null
        GraphScopes = @(
            "Organization.Read.All",
            "Directory.Read.All",
            "RoleManagement.Read.Directory",
            "User.Read.All"
        )
        GraphLicenseWriteScopes = @()
        EnableGraphAfterExchangeConnect = $true
    }

    if (-not [string]::IsNullOrWhiteSpace($Path)) {
        if (-not (Test-Path $Path -PathType Leaf)) {
            throw "Configuration file not found: $Path"
        }

        try {
            $document = Get-Content -Raw $Path | ConvertFrom-Json -Depth 20
        }
        catch {
            throw "Unable to parse configuration file '$Path': $($_.Exception.Message)"
        }

        $exchangeOnline = Get-ObjectPropertyValue -Object $document -PropertyName 'ExchangeOnline'
        if ($null -ne $exchangeOnline) {
            Merge-StringValue -Target $configuration -Key "AuthenticationMode" -Value (Get-ObjectPropertyValue -Object $exchangeOnline -PropertyName 'authenticationMode')
            Merge-StringValue -Target $configuration -Key "ExchangeEnvironmentName" -Value (Get-ObjectPropertyValue -Object $exchangeOnline -PropertyName 'exchangeEnvironmentName')
            Merge-StringValue -Target $configuration -Key "ExchangeOrganization" -Value (Get-ObjectPropertyValue -Object $exchangeOnline -PropertyName 'exchangeOrganization')
            Merge-StringValue -Target $configuration -Key "DelegatedOrganization" -Value (Get-ObjectPropertyValue -Object $exchangeOnline -PropertyName 'delegatedOrganization')
            Merge-StringValue -Target $configuration -Key "UserPrincipalNameHint" -Value (Get-ObjectPropertyValue -Object $exchangeOnline -PropertyName 'userPrincipalNameHint')
            Merge-StringValue -Target $configuration -Key "ApplicationId" -Value (Get-ObjectPropertyValue -Object $exchangeOnline -PropertyName 'applicationId')
            Merge-StringValue -Target $configuration -Key "CertificateThumbprint" -Value (Get-ObjectPropertyValue -Object $exchangeOnline -PropertyName 'certificateThumbprint')
            Merge-StringValue -Target $configuration -Key "CertificateSubjectName" -Value (Get-ObjectPropertyValue -Object $exchangeOnline -PropertyName 'certificateSubjectName')
            Merge-StringValue -Target $configuration -Key "ManagedIdentityAccountId" -Value (Get-ObjectPropertyValue -Object $exchangeOnline -PropertyName 'managedIdentityAccountId')
            Merge-StringValue -Target $configuration -Key "GraphTenantId" -Value (Get-ObjectPropertyValue -Object $exchangeOnline -PropertyName 'graphTenantId')
            Merge-StringListValue -Target $configuration -Key "GraphScopes" -Value (Get-ObjectPropertyValue -Object $exchangeOnline -PropertyName 'graphScopes')
            Merge-StringListValue -Target $configuration -Key "GraphLicenseWriteScopes" -Value (Get-ObjectPropertyValue -Object $exchangeOnline -PropertyName 'graphLicenseWriteScopes')
            $enableGraph = Read-OptionalBoolean (Get-ObjectPropertyValue -Object $exchangeOnline -PropertyName 'enableGraphAfterExchangeConnect')
            if ($null -ne $enableGraph) {
                $configuration.EnableGraphAfterExchangeConnect = $enableGraph
            }
        }
    }

    Merge-StringValue -Target $configuration -Key "AuthenticationMode" -Value (Get-EnvValue "ONLYEXO365_AUTH_MODE")
    Merge-StringValue -Target $configuration -Key "ExchangeEnvironmentName" -Value (Get-EnvValue "ONLYEXO365_EXO_ENV")
    Merge-StringValue -Target $configuration -Key "ExchangeOrganization" -Value (Get-EnvValue "ONLYEXO365_EXO_ORGANIZATION")
    Merge-StringValue -Target $configuration -Key "DelegatedOrganization" -Value (Get-EnvValue "ONLYEXO365_EXO_DELEGATED_ORGANIZATION")
    Merge-StringValue -Target $configuration -Key "UserPrincipalNameHint" -Value (Get-EnvValue "ONLYEXO365_EXO_UPN_HINT")
    Merge-StringValue -Target $configuration -Key "ApplicationId" -Value (Get-EnvValue "ONLYEXO365_APP_ID")
    Merge-StringValue -Target $configuration -Key "CertificateThumbprint" -Value (Get-EnvValue "ONLYEXO365_CERT_THUMBPRINT")
    Merge-StringValue -Target $configuration -Key "CertificateSubjectName" -Value (Get-EnvValue "ONLYEXO365_CERT_SUBJECT")
    Merge-StringValue -Target $configuration -Key "ManagedIdentityAccountId" -Value (Get-EnvValue "ONLYEXO365_MANAGED_IDENTITY_ACCOUNT_ID")
    Merge-StringValue -Target $configuration -Key "GraphTenantId" -Value (Get-EnvValue "ONLYEXO365_GRAPH_TENANT_ID")
    Merge-StringListValue -Target $configuration -Key "GraphScopes" -Value (Get-EnvValue "ONLYEXO365_GRAPH_SCOPES")
    Merge-StringListValue -Target $configuration -Key "GraphLicenseWriteScopes" -Value (Get-EnvValue "ONLYEXO365_GRAPH_LICENSE_WRITE_SCOPES")
    $enableGraphFromEnv = Read-OptionalBoolean (Get-EnvValue "ONLYEXO365_ENABLE_GRAPH")
    if ($null -ne $enableGraphFromEnv) {
        $configuration.EnableGraphAfterExchangeConnect = $enableGraphFromEnv
    }

    $configuration.AuthenticationMode = $configuration.AuthenticationMode.Trim()
    $configuration.GraphScopes = @($configuration.GraphScopes | Select-Object -Unique)
    $configuration.GraphLicenseWriteScopes = @($configuration.GraphLicenseWriteScopes | Select-Object -Unique)

    return $configuration
}

function Get-RequiredModulesFromPolicy {
    param([string]$PolicyPath)

    if (-not (Test-Path $PolicyPath -PathType Leaf)) {
        throw "Bootstrap policy not found: $PolicyPath"
    }

    try {
        $policy = Get-Content -Raw $PolicyPath | ConvertFrom-Json -Depth 20
    }
    catch {
        throw "Unable to parse bootstrap policy '$PolicyPath': $($_.Exception.Message)"
    }

    $requiredModules = New-Object System.Collections.Generic.List[object]
    $modules = @(Get-ObjectPropertyValue -Object $policy -PropertyName 'modules')
    if ($modules.Count -eq 0 -or $null -eq $modules[0]) { throw 'Bootstrap policy contains no module requirements.' }
    foreach ($module in $modules) {
        $version = Read-OptionalValue (Get-ObjectPropertyValue -Object $module -PropertyName 'requiredVersion')
        if ([string]::IsNullOrWhiteSpace($version)) {
            throw 'Bootstrap policy module missing requiredVersion.'
        }

        $moduleNames = @()
        $primaryModule = Read-OptionalValue (Get-ObjectPropertyValue -Object $module -PropertyName 'moduleName')
        if (-not [string]::IsNullOrWhiteSpace($primaryModule)) {
            $moduleNames += $primaryModule
        }

        foreach ($requiredModule in @(Read-StringList (Get-ObjectPropertyValue -Object $module -PropertyName 'requiredModules'))) {
            if ($moduleNames -notcontains $requiredModule) {
                $moduleNames += $requiredModule
            }
        }

        foreach ($moduleName in $moduleNames) {
            $requiredModules.Add([pscustomobject]@{
                module = $moduleName
                required_version = $version
            })
        }
        if ($moduleNames.Count -eq 0) { throw 'Bootstrap policy module missing moduleName and requiredModules.' }
    }

    return @(
        $requiredModules |
            Group-Object module |
            ForEach-Object {
                $versions = @($_.Group.required_version | Select-Object -Unique)
                if ($versions.Count -ne 1) {
                    throw "Bootstrap policy defines multiple versions for module $($_.Name): $($versions -join ', ')"
                }

                [pscustomobject]@{
                    module = $_.Name
                    required_version = $versions[0]
                }
            }
    )
}

function Ensure-RequiredModule {
    param(
        [string]$ModuleName,
        [string]$RequiredVersion,
        [bool]$InstallIfMissing
    )

    $installed = Get-Module -ListAvailable -Name $ModuleName |
        Where-Object { $_.Version -eq [Version]$RequiredVersion } |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if ($null -ne $installed) {
        Write-Success "$ModuleName $RequiredVersion available."
        return [ordered]@{
            module = $ModuleName
            required_version = $RequiredVersion
            installed = $true
            installed_version = $installed.Version.ToString()
        }
    }

    if (-not $InstallIfMissing) {
        throw "Required module $ModuleName $RequiredVersion is not installed."
    }

    Write-Info "Installing approved module $ModuleName $RequiredVersion for CurrentUser."
    $installPsResource = Get-Command Install-PSResource -ErrorAction SilentlyContinue
    if ($null -ne $installPsResource) {
        $arguments = @{
            Name = $ModuleName
            Version = $RequiredVersion
            Repository = "PSGallery"
            Scope = "CurrentUser"
            TrustRepository = $true
            Quiet = $true
            ErrorAction = "Stop"
        }

        if ($installPsResource.Parameters.ContainsKey("AcceptLicense")) {
            $arguments.AcceptLicense = $true
        }

        Install-PSResource @arguments
    }
    else {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -ForceBootstrap -Force -Scope CurrentUser -Confirm:$false -ErrorAction Stop | Out-Null
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue

        $installModuleCommand = Get-Command Install-Module -ErrorAction Stop
        $arguments = @{
            Name = $ModuleName
            Repository = "PSGallery"
            RequiredVersion = $RequiredVersion
            Scope = "CurrentUser"
            Force = $true
            AllowClobber = $true
            Confirm = $false
            ErrorAction = "Stop"
        }

        if ($installModuleCommand.Parameters.ContainsKey("AcceptLicense")) {
            $arguments.AcceptLicense = $true
        }

        Install-Module @arguments
    }

    $installed = Get-Module -ListAvailable -Name $ModuleName |
        Where-Object { $_.Version -eq [Version]$RequiredVersion } |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if ($null -eq $installed) {
        throw "Module installation completed but $ModuleName $RequiredVersion is still unavailable."
    }

    Write-Success "$ModuleName $RequiredVersion installed."
    return [ordered]@{
        module = $ModuleName
        required_version = $RequiredVersion
        installed = $true
        installed_version = $installed.Version.ToString()
    }
}

function Format-CommandPreviewValue {
    param([object]$Value)

    if ($Value -is [bool]) {
        if ($Value) {
            return "`$true"
        }

        return "`$false"
    }

    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $items = @($Value | ForEach-Object { "'$(([string]$_).Replace("'", "''"))'" })
        return "@($($items -join ', '))"
    }

    return "'$(([string]$Value).Replace("'", "''"))'"
}

function New-CommandInvocation {
    param(
        [string]$CommandName,
        [System.Collections.IDictionary]$Arguments
    )

    $previewParts = New-Object System.Collections.Generic.List[string]
    $previewParts.Add($CommandName)

    foreach ($entry in $Arguments.GetEnumerator()) {
        if ($entry.Value -is [bool] -and $entry.Value) {
            $previewParts.Add("-$($entry.Key)")
            continue
        }

        $previewParts.Add("-$($entry.Key)")
        $previewParts.Add((Format-CommandPreviewValue -Value $entry.Value))
    }

    return [pscustomobject]@{
        CommandName = $CommandName
        Arguments = $Arguments
        Preview = $previewParts -join " "
    }
}

function Add-OptionalCommandArgument {
    param(
        [System.Collections.IDictionary]$Arguments,
        [string]$Name,
        [object]$Value
    )

    if (-not [string]::IsNullOrWhiteSpace([string]$Value)) {
        $Arguments[$Name] = $Value
    }
}

function Get-ExchangeConnectCommand {
    param(
        [hashtable]$Configuration,
        [string]$AuthenticationMode
    )

    $arguments = [ordered]@{
        ShowBanner = $false
    }

    Add-OptionalCommandArgument -Arguments $arguments -Name "ExchangeEnvironmentName" -Value $Configuration.ExchangeEnvironmentName
    Add-OptionalCommandArgument -Arguments $arguments -Name "DelegatedOrganization" -Value $Configuration.DelegatedOrganization

    switch ($AuthenticationMode.ToLowerInvariant()) {
        "interactive" {
            Add-OptionalCommandArgument -Arguments $arguments -Name "Organization" -Value $Configuration.ExchangeOrganization
            Add-OptionalCommandArgument -Arguments $arguments -Name "UserPrincipalName" -Value $Configuration.UserPrincipalNameHint
            return New-CommandInvocation -CommandName "Connect-ExchangeOnline" -Arguments $arguments
        }
        "devicecode" {
            Add-OptionalCommandArgument -Arguments $arguments -Name "Organization" -Value $Configuration.ExchangeOrganization
            Add-OptionalCommandArgument -Arguments $arguments -Name "UserPrincipalName" -Value $Configuration.UserPrincipalNameHint
            $arguments.Device = $true
            return New-CommandInvocation -CommandName "Connect-ExchangeOnline" -Arguments $arguments
        }
        "appcertificate" {
            if ([string]::IsNullOrWhiteSpace($Configuration.ApplicationId) -or [string]::IsNullOrWhiteSpace($Configuration.ExchangeOrganization)) {
                throw "AppCertificate validation requires ONLYEXO365_APP_ID and ONLYEXO365_EXO_ORGANIZATION."
            }

            $arguments.AppId = $Configuration.ApplicationId
            $arguments.Organization = $Configuration.ExchangeOrganization
            if (-not [string]::IsNullOrWhiteSpace($Configuration.CertificateThumbprint)) {
                $arguments.CertificateThumbprint = $Configuration.CertificateThumbprint
            }
            elseif (-not [string]::IsNullOrWhiteSpace($Configuration.CertificateSubjectName)) {
                $arguments.CertificateSubjectName = $Configuration.CertificateSubjectName
            }
            else {
                throw "AppCertificate validation requires ONLYEXO365_CERT_THUMBPRINT or ONLYEXO365_CERT_SUBJECT."
            }

            return New-CommandInvocation -CommandName "Connect-ExchangeOnline" -Arguments $arguments
        }
        "managedidentity" {
            if ([string]::IsNullOrWhiteSpace($Configuration.ExchangeOrganization)) {
                throw "ManagedIdentity validation requires ONLYEXO365_EXO_ORGANIZATION."
            }

            $arguments.ManagedIdentity = $true
            $arguments.Organization = $Configuration.ExchangeOrganization
            Add-OptionalCommandArgument -Arguments $arguments -Name "ManagedIdentityAccountId" -Value $Configuration.ManagedIdentityAccountId
            return New-CommandInvocation -CommandName "Connect-ExchangeOnline" -Arguments $arguments
        }
        default {
            throw "Unsupported authentication mode for tenant validation: $($Configuration.AuthenticationMode)"
        }
    }
}

function Get-GraphConnectCommand {
    param(
        [hashtable]$Configuration,
        [string]$AuthenticationMode
    )

    $scopes = @($Configuration.GraphScopes)
    if ($scopes.Count -eq 0) {
        $scopes = @("Organization.Read.All")
    }

    switch ($AuthenticationMode.ToLowerInvariant()) {
        "interactive" {
            $arguments = [ordered]@{
                Scopes = $scopes
                ContextScope = "Process"
                NoWelcome = $true
            }
            Add-OptionalCommandArgument -Arguments $arguments -Name "TenantId" -Value $Configuration.GraphTenantId
            return New-CommandInvocation -CommandName "Connect-MgGraph" -Arguments $arguments
        }
        "devicecode" {
            $arguments = [ordered]@{
                Scopes = $scopes
                ContextScope = "Process"
                UseDeviceCode = $true
                NoWelcome = $true
            }
            Add-OptionalCommandArgument -Arguments $arguments -Name "TenantId" -Value $Configuration.GraphTenantId
            return New-CommandInvocation -CommandName "Connect-MgGraph" -Arguments $arguments
        }
        "appcertificate" {
            if ([string]::IsNullOrWhiteSpace($Configuration.ApplicationId) -or [string]::IsNullOrWhiteSpace($Configuration.GraphTenantId)) {
                throw "AppCertificate validation requires ONLYEXO365_APP_ID and ONLYEXO365_GRAPH_TENANT_ID for Graph."
            }

            $arguments = [ordered]@{
                ClientId = $Configuration.ApplicationId
                TenantId = $Configuration.GraphTenantId
                ContextScope = "Process"
                NoWelcome = $true
            }
            if (-not [string]::IsNullOrWhiteSpace($Configuration.CertificateThumbprint)) {
                $arguments.CertificateThumbprint = $Configuration.CertificateThumbprint
            }
            elseif (-not [string]::IsNullOrWhiteSpace($Configuration.CertificateSubjectName)) {
                $arguments.CertificateSubjectName = $Configuration.CertificateSubjectName
            }
            else {
                throw "AppCertificate validation requires ONLYEXO365_CERT_THUMBPRINT or ONLYEXO365_CERT_SUBJECT for Graph."
            }

            return New-CommandInvocation -CommandName "Connect-MgGraph" -Arguments $arguments
        }
        "managedidentity" {
            $arguments = [ordered]@{
                Identity = $true
                ContextScope = "Process"
                NoWelcome = $true
            }
            Add-OptionalCommandArgument -Arguments $arguments -Name "ClientId" -Value $Configuration.ManagedIdentityAccountId
            return New-CommandInvocation -CommandName "Connect-MgGraph" -Arguments $arguments
        }
        default {
            throw "Unsupported authentication mode for Graph validation: $($Configuration.AuthenticationMode)"
        }
    }
}

function Get-ComplianceConnectCommand {
    param([hashtable]$Configuration, [string]$AuthenticationMode, [switch]$SearchOnly)
    $arguments = [ordered]@{ ShowBanner = $false }
    $authority = 'https://login.microsoftonline.com'
    switch ($Configuration.ExchangeEnvironmentName) {
        'O365Default' { }
        'O365USGovGCCHigh' {
            $arguments.ConnectionUri = 'https://ps.compliance.protection.office365.us/powershell-liveid/'
            $authority = 'https://login.microsoftonline.us'
        }
        'O365USGovDoD' {
            $arguments.ConnectionUri = 'https://l5.ps.compliance.protection.office365.us/powershell-liveid/'
            $authority = 'https://login.microsoftonline.us'
        }
        'O365China' {
            $arguments.ConnectionUri = 'https://ps.compliance.protection.partner.outlook.cn/powershell-liveid'
            $authority = 'https://login.chinacloudapi.cn'
        }
        default { throw "Unsupported Compliance cloud: $($Configuration.ExchangeEnvironmentName)" }
    }
    $tenant = if ([string]::IsNullOrWhiteSpace($Configuration.DelegatedOrganization)) { 'organizations' } else { [uri]::EscapeDataString($Configuration.DelegatedOrganization) }
    $arguments.AzureADAuthorizationEndpointUri = "$authority/$tenant"
    Add-OptionalCommandArgument -Arguments $arguments -Name 'DelegatedOrganization' -Value $Configuration.DelegatedOrganization
    switch ($AuthenticationMode.ToLowerInvariant()) {
        'interactive' {
            Add-OptionalCommandArgument -Arguments $arguments -Name 'UserPrincipalName' -Value $Configuration.UserPrincipalNameHint
        }
        'appcertificate' {
            if ($SearchOnly) { throw 'eDiscovery requires delegated interactive authentication; app-only authentication is unsupported.' }
            $exchange = Get-ExchangeConnectCommand -Configuration $Configuration -AuthenticationMode $AuthenticationMode
            foreach ($name in @('AppId', 'Organization', 'CertificateThumbprint', 'CertificateSubjectName')) {
                if ($exchange.Arguments.Contains($name)) { $arguments[$name] = $exchange.Arguments[$name] }
            }
            if ($arguments.Contains('CertificateSubjectName')) {
                $certificates = @(Get-ChildItem -LiteralPath 'Cert:\CurrentUser\My' | Where-Object Subject -eq $arguments.CertificateSubjectName)
                if ($certificates.Count -ne 1) { throw 'Compliance certificate subject must identify exactly one certificate.' }
                $arguments.Remove('CertificateSubjectName')
                $arguments.CertificateThumbprint = $certificates[0].Thumbprint
            }
        }
        default { throw "Security & Compliance does not support the configured authentication mode '$AuthenticationMode'. Use Interactive or AppCertificate; eDiscovery requires Interactive." }
    }
    if ($SearchOnly) { $arguments.EnableSearchOnlySession = $true }
    return New-CommandInvocation -CommandName 'Connect-IPPSSession' -Arguments $arguments
}

function Invoke-Probe {
    param(
        [string]$Name,
        [scriptblock]$ScriptBlock
    )

    Write-Info "Running probe: $Name"
    try {
        $data = & $ScriptBlock
        $result = [ordered]@{
            success = $true
            error = $null
        }

        if ($data -is [System.Collections.IDictionary]) {
            foreach ($entry in $data.GetEnumerator()) {
                $result[$entry.Key] = $entry.Value
            }
        }
        elseif ($null -ne $data) {
            $result.data = $data
        }

        Write-Success "Probe completed: $Name"
        return $result
    }
    catch {
        $message = $_.Exception.Message
        Write-Warn "Probe failed: $Name. $message"
        return [ordered]@{
            success = $false
            error = $message
        }
    }
}

function Invoke-CommandAvailabilityProbe {
    param(
        [string]$CommandName,
        [string[]]$RequiredParameters = @(),
        [bool]$SessionConnected = $true
    )

    return Invoke-Probe -Name "Get-Command $CommandName" -ScriptBlock {
        if (-not $SessionConnected) { throw 'Required Security & Compliance session was not established.' }
        $command = Get-Command -Name $CommandName -ErrorAction Stop
        $parameters = @($command.Parameters.Keys | Sort-Object)
        $missingRequiredParameters = @($RequiredParameters | Where-Object { $parameters -notcontains $_ })

        if ($missingRequiredParameters.Count -gt 0) {
            throw "Cmdlet $CommandName missing required parameters: $($missingRequiredParameters -join ', ')"
        }

        [ordered]@{
            command_name = $CommandName
            module_name = [string]$command.ModuleName
            available = $true
            parameter_count = $parameters.Count
            required_parameters = @($RequiredParameters)
        }
    }
}

function Get-ReportProbeResults {
    param([hashtable]$Report)

    $results = New-Object System.Collections.Generic.List[object]

    foreach ($sectionName in @("exchange", "compliance", "mail_security", "graph")) {
        $section = $Report[$sectionName]
        if ($null -eq $section -or $null -eq $section.probes) {
            continue
        }

        foreach ($probe in $section.probes.GetEnumerator()) {
            $results.Add([pscustomobject]@{
                section = $sectionName
                probe = [string]$probe.Key
                success = [bool]$probe.Value.success
                error = [string]$probe.Value.error
            })
        }
    }

    return $results.ToArray()
}

$resolvedConfigurationPath = if ([string]::IsNullOrWhiteSpace($ConfigurationPath)) { $null } else { Resolve-RepositoryPath -RepositoryRoot $repositoryRoot -PathValue $ConfigurationPath }
$resolvedReportPath = Resolve-RepositoryPath -RepositoryRoot $repositoryRoot -PathValue $ReportPath
$reportDirectory = Split-Path -Parent $resolvedReportPath
$bootstrapPolicyPath = Join-Path $repositoryRoot "src\OnlyExo365.Contracts\Data\PowerShellModuleBootstrapPolicy.json"

if (-not (Test-Path $reportDirectory)) {
    New-Item -Path $reportDirectory -ItemType Directory -Force | Out-Null
}

$report = [ordered]@{
    generated_at_utc = (Get-Date).ToUniversalTime().ToString("O")
    configuration = [ordered]@{}
    prerequisites = [ordered]@{
        pwsh_major_version = $PSVersionTable.PSVersion.Major
        required_modules = @()
    }
    exchange = [ordered]@{
        connected = $false
        probes = [ordered]@{}
    }
    compliance = [ordered]@{
        attempted = $false
        exchange_session_available = $false
        connected = $false
        search_session_connected = $false
        probes = [ordered]@{}
    }
    mail_security = [ordered]@{
        attempted = $false
        exchange_session_available = $false
        probes = [ordered]@{}
    }
    graph = [ordered]@{
        attempted = (-not [bool]$SkipGraph)
        connected = $false
        probes = [ordered]@{}
    }
    summary = [ordered]@{
        total_probe_count = 0
        failed_probe_count = 0
        failed_probes = @()
    }
}

$failedProbes = @()

try {
    $configuration = Load-ExchangeConfiguration -Path $resolvedConfigurationPath
    $authResolution = Resolve-EffectiveAuthenticationMode -Configuration $configuration
    $requiredModulesFromPolicy = Get-RequiredModulesFromPolicy -PolicyPath $bootstrapPolicyPath

    $report.configuration = [ordered]@{
        requested_authentication_mode = $authResolution.requested_mode
        effective_authentication_mode = $authResolution.effective_mode
        used_device_code_fallback = [bool]$authResolution.used_device_code_fallback
        authentication_fallback_reason = $authResolution.fallback_reason
        exchange_environment_name = $configuration.ExchangeEnvironmentName
        exchange_organization = $configuration.ExchangeOrganization
        delegated_organization = $configuration.DelegatedOrganization
        graph_tenant_id = $configuration.GraphTenantId
        enable_graph_after_exchange_connect = [bool]$configuration.EnableGraphAfterExchangeConnect
        graph_validation_requested = -not [bool]$SkipGraph
        graph_scopes = @($configuration.GraphScopes)
    }
    Write-Step "Ensuring approved PowerShell modules"
    $requiredModules = New-Object System.Collections.Generic.List[object]
    foreach ($moduleRequirement in $requiredModulesFromPolicy) {
        $moduleName = [string]$moduleRequirement.module
        if ($SkipGraph -and $moduleName.StartsWith("Microsoft.Graph.", [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }

        $requiredVersion = [string]$moduleRequirement.required_version
        if ([string]::IsNullOrWhiteSpace($requiredVersion)) {
            throw "Bootstrap policy missing required version for $moduleName."
        }

        $requiredModules.Add((Ensure-RequiredModule -ModuleName $moduleName -RequiredVersion $requiredVersion -InstallIfMissing ([bool]$InstallApprovedModules)))
        Import-Module -FullyQualifiedName @{ ModuleName = $moduleName; RequiredVersion = $requiredVersion } -Global -ErrorAction Stop
    }

    $report.prerequisites.required_modules = $requiredModules.ToArray()

    Write-Step "Connecting to Exchange Online"
    $exchangeConnectCommand = Get-ExchangeConnectCommand -Configuration $configuration -AuthenticationMode $authResolution.effective_mode
    Write-Info $exchangeConnectCommand.Preview
    $exchangeConnectArguments = $exchangeConnectCommand.Arguments
    & $exchangeConnectCommand.CommandName @exchangeConnectArguments | Out-Null
    $report.exchange.connected = $true
    Write-Success "Exchange Online session established."

    $report.exchange.probes.organization = Invoke-Probe -Name "Get-OrganizationConfig" -ScriptBlock {
        $organization = Get-OrganizationConfig -ErrorAction Stop
        [ordered]@{
            name = $organization.Name
            public_folders_enabled = [bool]$organization.PublicFoldersEnabled
        }
    }

    $report.exchange.probes.accepted_domains = Invoke-Probe -Name "Get-AcceptedDomain" -ScriptBlock {
        $domains = @(Get-AcceptedDomain -ErrorAction Stop | Sort-Object DomainName | Select-Object -First 5)
        [ordered]@{
            sample_count = $domains.Count
            sample = @($domains | ForEach-Object {
                [ordered]@{
                    domain_name = [string]$_.DomainName
                    domain_type = [string]$_.DomainType
                    is_default = [bool]$_.Default
                }
            })
        }
    }

    $report.exchange.probes.mailboxes = Invoke-Probe -Name "Get-EXOMailbox" -ScriptBlock {
        $mailboxes = @(Get-EXOMailbox -ResultSize 3 -Properties RecipientTypeDetails,PrimarySmtpAddress -ErrorAction Stop)
        [ordered]@{
            sample_count = $mailboxes.Count
            sample = @($mailboxes | ForEach-Object {
                [ordered]@{
                    identity = [string]$_.ExternalDirectoryObjectId
                    primary_smtp_address = [string]$_.PrimarySmtpAddress
                    recipient_type_details = [string]$_.RecipientTypeDetails
                }
            })
        }
    }

    Write-Step "Running Compliance validation probes"
    $report.compliance.attempted = $true
    $report.compliance.exchange_session_available = $true
    $report.compliance.probes.connection = Invoke-Probe -Name 'Connect-IPPSSession' -ScriptBlock {
        $command = Get-ComplianceConnectCommand -Configuration $configuration -AuthenticationMode $authResolution.effective_mode
        $arguments = $command.Arguments
        & $command.CommandName @arguments | Out-Null
        $report.compliance.connected = $true
        return [ordered]@{ connected = $true }
    }

    $report.compliance.probes.search_unified_audit_log = Invoke-Probe -Name "Search-UnifiedAuditLog" -ScriptBlock {
        $endDate = Get-Date
        $startDate = $endDate.AddDays(-1)
        $records = @(Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate -ResultSize 1 -ErrorAction Stop | Sort-Object CreationDate -Descending | Select-Object -First 1)
        $latestRecord = $records | Select-Object -First 1
        [ordered]@{
            sample_count = $records.Count
            sampled_window_start_utc = $startDate.ToUniversalTime().ToString("O")
            sampled_window_end_utc = $endDate.ToUniversalTime().ToString("O")
            sample = @($records | ForEach-Object {
                [ordered]@{
                    identity = if ($_.Identity) { $_.Identity.ToString() } else { $null }
                    creation_date = if ($_.CreationDate) { $_.CreationDate } else { $null }
                    operations = if ($_.Operations) { $_.Operations.ToString() } else { $null }
                    result_status = if ($_.ResultStatus) { $_.ResultStatus.ToString() } else { $null }
                }
            })
            latest_record_identity = if ($null -ne $latestRecord -and $latestRecord.Identity) { $latestRecord.Identity.ToString() } else { $null }
        }
    }

    $report.compliance.probes.new_case_hold_policy = Invoke-CommandAvailabilityProbe -CommandName "New-CaseHoldPolicy" -RequiredParameters @("Name", "Case", "ExchangeLocation") -SessionConnected $report.compliance.connected
    $report.compliance.probes.new_case_hold_rule = Invoke-CommandAvailabilityProbe -CommandName "New-CaseHoldRule" -RequiredParameters @("Name", "Policy", "ContentMatchQuery") -SessionConnected $report.compliance.connected

    $report.compliance.probes.search_connection = Invoke-Probe -Name 'Connect-IPPSSession -EnableSearchOnlySession' -ScriptBlock {
        $command = Get-ComplianceConnectCommand -Configuration $configuration -AuthenticationMode $authResolution.effective_mode -SearchOnly
        $arguments = $command.Arguments
        & $command.CommandName @arguments | Out-Null
        $report.compliance.search_session_connected = $true
        return [ordered]@{ connected = $true }
    }
    $report.compliance.probes.compliance_searches = Invoke-Probe -Name "Get-ComplianceSearch" -ScriptBlock {
        if (-not $report.compliance.search_session_connected) { throw 'eDiscovery search session was not established.' }
        $searches = @(Get-ComplianceSearch -ErrorAction Stop | Sort-Object Name | Select-Object -First 3)
        [ordered]@{
            sample_count = $searches.Count
            sample = @($searches | ForEach-Object {
                [ordered]@{
                    name = if ($_.Name) { $_.Name.ToString() } else { $null }
                    status = if ($_.Status) { $_.Status.ToString() } else { $null }
                    case = if ($_.Case) { $_.Case.ToString() } else { $null }
                }
            })
        }
    }

    $report.compliance.probes.new_compliance_search = Invoke-CommandAvailabilityProbe -CommandName "New-ComplianceSearch" -RequiredParameters @("Name", "ExchangeLocation") -SessionConnected $report.compliance.search_session_connected
    $report.compliance.probes.start_compliance_search = Invoke-CommandAvailabilityProbe -CommandName "Start-ComplianceSearch" -RequiredParameters @("Identity") -SessionConnected $report.compliance.search_session_connected
    $report.compliance.probes.remove_compliance_search = Invoke-CommandAvailabilityProbe -CommandName "Remove-ComplianceSearch" -RequiredParameters @("Identity") -SessionConnected $report.compliance.search_session_connected
    $report.compliance.probes.new_compliance_search_action = Invoke-CommandAvailabilityProbe -CommandName "New-ComplianceSearchAction" -RequiredParameters @("SearchName", "Purge", "PurgeType") -SessionConnected $report.compliance.search_session_connected

    Write-Step "Running Mail Security validation probes"
    $report.mail_security.attempted = $true
    $report.mail_security.exchange_session_available = $true

    $report.mail_security.probes.dkim_signing = Invoke-Probe -Name "Get-DkimSigningConfig" -ScriptBlock {
        $configs = @(Get-DkimSigningConfig -ErrorAction Stop | Sort-Object Identity | Select-Object -First 3)
        [ordered]@{
            sample_count = $configs.Count
            sample = @($configs | ForEach-Object {
                [ordered]@{
                    identity = if ($_.Identity) { $_.Identity.ToString() } else { $null }
                    domain = if ($_.Domain) { $_.Domain.ToString() } else { $null }
                }
            })
        }
    }

    $report.mail_security.probes.hosted_content_filter_policies = Invoke-Probe -Name "Get-HostedContentFilterPolicy" -ScriptBlock {
        $policies = @(Get-HostedContentFilterPolicy -ErrorAction Stop | Sort-Object Name | Select-Object -First 3)
        [ordered]@{
            sample_count = $policies.Count
            sample = @($policies | ForEach-Object {
                [ordered]@{
                    identity = if ($_.Identity) { $_.Identity.ToString() } else { $null }
                    name = if ($_.Name) { $_.Name.ToString() } else { $null }
                }
            })
        }
    }

    $report.mail_security.probes.anti_phish_policies = Invoke-Probe -Name "Get-AntiPhishPolicy" -ScriptBlock {
        $policies = @(Get-AntiPhishPolicy -ErrorAction Stop | Sort-Object Name | Select-Object -First 3)
        [ordered]@{
            sample_count = $policies.Count
            sample = @($policies | ForEach-Object {
                [ordered]@{
                    identity = if ($_.Identity) { $_.Identity.ToString() } else { $null }
                    name = if ($_.Name) { $_.Name.ToString() } else { $null }
                }
            })
        }
    }

    $report.mail_security.probes.malware_filter_policies = Invoke-Probe -Name "Get-MalwareFilterPolicy" -ScriptBlock {
        $policies = @(Get-MalwareFilterPolicy -ErrorAction Stop | Sort-Object Name | Select-Object -First 3)
        [ordered]@{
            sample_count = $policies.Count
            sample = @($policies | ForEach-Object {
                [ordered]@{
                    identity = if ($_.Identity) { $_.Identity.ToString() } else { $null }
                    name = if ($_.Name) { $_.Name.ToString() } else { $null }
                }
            })
        }
    }

    $report.mail_security.probes.quarantine_policies = Invoke-Probe -Name "Get-QuarantinePolicy" -ScriptBlock {
        $policies = @(Get-QuarantinePolicy -ErrorAction Stop | Sort-Object Name | Select-Object -First 3)
        [ordered]@{
            sample_count = $policies.Count
            sample = @($policies | ForEach-Object {
                [ordered]@{
                    identity = if ($_.Identity) { $_.Identity.ToString() } else { $null }
                    name = if ($_.Name) { $_.Name.ToString() } else { $null }
                }
            })
        }
    }

    $report.mail_security.probes.hosted_outbound_spam_filter_policies = Invoke-Probe -Name "Get-HostedOutboundSpamFilterPolicy" -ScriptBlock {
        $policies = @(Get-HostedOutboundSpamFilterPolicy -ErrorAction Stop | Sort-Object Name | Select-Object -First 3)
        [ordered]@{
            sample_count = $policies.Count
            sample = @($policies | ForEach-Object {
                [ordered]@{
                    identity = if ($_.Identity) { $_.Identity.ToString() } else { $null }
                    name = if ($_.Name) { $_.Name.ToString() } else { $null }
                }
            })
        }
    }

    if (-not $SkipGraph) {
        Write-Step "Connecting to Microsoft Graph"
        $graphConnectCommand = Get-GraphConnectCommand -Configuration $configuration -AuthenticationMode $authResolution.effective_mode
        Write-Info $graphConnectCommand.Preview
        $graphConnectArguments = $graphConnectCommand.Arguments
        & $graphConnectCommand.CommandName @graphConnectArguments | Out-Null
        $report.graph.connected = $true
        Write-Success "Microsoft Graph session established."

        $report.graph.probes.organization = Invoke-Probe -Name "Get-MgOrganization" -ScriptBlock {
            $organization = @(Get-MgOrganization -ErrorAction Stop | Select-Object -First 1)
            if ($organization.Count -eq 0) {
                throw "No organization returned by Graph."
            }

            [ordered]@{
                id = [string]$organization[0].Id
                display_name = [string]$organization[0].DisplayName
            }
        }

        $report.graph.probes.subscribed_skus = Invoke-Probe -Name "Get-MgSubscribedSku" -ScriptBlock {
            $skus = @(Get-MgSubscribedSku -ErrorAction Stop | Sort-Object SkuPartNumber | Select-Object -First 5)
            [ordered]@{
                sample_count = $skus.Count
                sample = @($skus | ForEach-Object {
                    [ordered]@{
                        sku_part_number = [string]$_.SkuPartNumber
                        sku_id = [string]$_.SkuId
                        enabled_units = if ($null -ne $_.PrepaidUnits) { [int]$_.PrepaidUnits.Enabled } else { 0 }
                        consumed_units = [int]$_.ConsumedUnits
                    }
                })
            }
        }
    }
    else {
        Write-Info "Graph validation skipped by request."
    }

    $probeResults = Get-ReportProbeResults -Report $report
    $failedProbes = @($probeResults | Where-Object { -not $_.success })
    $report.summary.total_probe_count = $probeResults.Count
    $report.summary.failed_probe_count = $failedProbes.Count
    $report.summary.failed_probes = @($failedProbes | ForEach-Object {
        [ordered]@{
            section = $_.section
            probe = $_.probe
            error = $_.error
        }
    })
}
catch {
    $report.summary.fatal_error = $_.Exception.Message
    throw
}
finally {
    try {
        $report | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $resolvedReportPath -Encoding utf8NoBOM
        Write-Info "Tenant validation report: $resolvedReportPath"
    }
    finally {
        try {
            if ($report.exchange.connected -or $report.compliance.connected) {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop | Out-Null
            }
        }
        finally {
            if ($report.graph.connected) { Disconnect-MgGraph -ErrorAction Stop | Out-Null }
        }
    }
}

if ($failedProbes.Count -gt 0) {
    throw "Tenant validation completed with $($failedProbes.Count) failed probe(s). See report: $resolvedReportPath"
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " TENANT VALIDATION COMPLETED" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

