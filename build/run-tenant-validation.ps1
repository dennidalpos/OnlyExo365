#Requires -Version 7.0

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

. (Join-Path $scriptDirectory "helpers\common.ps1")

function Resolve-RepoPath {
    param(
        [string]$BaseDirectory,
        [string]$PathValue
    )

    if ([string]::IsNullOrWhiteSpace($PathValue)) {
        return $PathValue
    }

    if ([System.IO.Path]::IsPathRooted($PathValue)) {
        return [System.IO.Path]::GetFullPath($PathValue)
    }

    return [System.IO.Path]::GetFullPath((Join-Path $BaseDirectory $PathValue))
}

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
            Stop-WithError "Configuration file not found: $Path"
        }

        try {
            $document = Get-Content -Raw $Path | ConvertFrom-Json -Depth 20
        }
        catch {
            Stop-WithError "Unable to parse configuration file '$Path': $($_.Exception.Message)"
        }

        $exchangeOnline = $document.ExchangeOnline
        if ($null -ne $exchangeOnline) {
            Merge-StringValue -Target $configuration -Key "AuthenticationMode" -Value $exchangeOnline.authenticationMode
            Merge-StringValue -Target $configuration -Key "ExchangeEnvironmentName" -Value $exchangeOnline.exchangeEnvironmentName
            Merge-StringValue -Target $configuration -Key "ExchangeOrganization" -Value $exchangeOnline.exchangeOrganization
            Merge-StringValue -Target $configuration -Key "DelegatedOrganization" -Value $exchangeOnline.delegatedOrganization
            Merge-StringValue -Target $configuration -Key "UserPrincipalNameHint" -Value $exchangeOnline.userPrincipalNameHint
            Merge-StringValue -Target $configuration -Key "ApplicationId" -Value $exchangeOnline.applicationId
            Merge-StringValue -Target $configuration -Key "CertificateThumbprint" -Value $exchangeOnline.certificateThumbprint
            Merge-StringValue -Target $configuration -Key "CertificateSubjectName" -Value $exchangeOnline.certificateSubjectName
            Merge-StringValue -Target $configuration -Key "ManagedIdentityAccountId" -Value $exchangeOnline.managedIdentityAccountId
            Merge-StringValue -Target $configuration -Key "GraphTenantId" -Value $exchangeOnline.graphTenantId
            Merge-StringListValue -Target $configuration -Key "GraphScopes" -Value $exchangeOnline.graphScopes
            Merge-StringListValue -Target $configuration -Key "GraphLicenseWriteScopes" -Value $exchangeOnline.graphLicenseWriteScopes
            $enableGraph = Read-OptionalBoolean $exchangeOnline.enableGraphAfterExchangeConnect
            if ($null -ne $enableGraph) {
                $configuration.EnableGraphAfterExchangeConnect = $enableGraph
            }
        }
    }

    Merge-StringValue -Target $configuration -Key "AuthenticationMode" -Value (Get-EnvValue "EXCHANGEADMIN_AUTH_MODE")
    Merge-StringValue -Target $configuration -Key "ExchangeEnvironmentName" -Value (Get-EnvValue "EXCHANGEADMIN_EXO_ENV")
    Merge-StringValue -Target $configuration -Key "ExchangeOrganization" -Value (Get-EnvValue "EXCHANGEADMIN_EXO_ORGANIZATION")
    Merge-StringValue -Target $configuration -Key "DelegatedOrganization" -Value (Get-EnvValue "EXCHANGEADMIN_EXO_DELEGATED_ORGANIZATION")
    Merge-StringValue -Target $configuration -Key "UserPrincipalNameHint" -Value (Get-EnvValue "EXCHANGEADMIN_EXO_UPN_HINT")
    Merge-StringValue -Target $configuration -Key "ApplicationId" -Value (Get-EnvValue "EXCHANGEADMIN_APP_ID")
    Merge-StringValue -Target $configuration -Key "CertificateThumbprint" -Value (Get-EnvValue "EXCHANGEADMIN_CERT_THUMBPRINT")
    Merge-StringValue -Target $configuration -Key "CertificateSubjectName" -Value (Get-EnvValue "EXCHANGEADMIN_CERT_SUBJECT")
    Merge-StringValue -Target $configuration -Key "ManagedIdentityAccountId" -Value (Get-EnvValue "EXCHANGEADMIN_MANAGED_IDENTITY_ACCOUNT_ID")
    Merge-StringValue -Target $configuration -Key "GraphTenantId" -Value (Get-EnvValue "EXCHANGEADMIN_GRAPH_TENANT_ID")
    Merge-StringListValue -Target $configuration -Key "GraphScopes" -Value (Get-EnvValue "EXCHANGEADMIN_GRAPH_SCOPES")
    Merge-StringListValue -Target $configuration -Key "GraphLicenseWriteScopes" -Value (Get-EnvValue "EXCHANGEADMIN_GRAPH_LICENSE_WRITE_SCOPES")
    $enableGraphFromEnv = Read-OptionalBoolean (Get-EnvValue "EXCHANGEADMIN_ENABLE_GRAPH")
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
        Stop-WithError "Bootstrap policy not found: $PolicyPath"
    }

    try {
        $policy = Get-Content -Raw $PolicyPath | ConvertFrom-Json -Depth 20
    }
    catch {
        Stop-WithError "Unable to parse bootstrap policy '$PolicyPath': $($_.Exception.Message)"
    }

    $requiredModules = New-Object System.Collections.Generic.List[object]
    foreach ($module in @($policy.modules)) {
        $version = Read-OptionalValue $module.requiredVersion
        if ([string]::IsNullOrWhiteSpace($version)) {
            continue
        }

        $moduleNames = @()
        $primaryModule = Read-OptionalValue $module.moduleName
        if (-not [string]::IsNullOrWhiteSpace($primaryModule)) {
            $moduleNames += $primaryModule
        }

        foreach ($requiredModule in @(Read-StringList $module.requiredModules)) {
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
    }

    return @(
        $requiredModules |
            Group-Object module |
            ForEach-Object {
                $versions = @($_.Group.required_version | Select-Object -Unique)
                if ($versions.Count -ne 1) {
                    Stop-WithError "Bootstrap policy defines multiple versions for module $($_.Name): $($versions -join ', ')"
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
        Stop-WithError "Required module $ModuleName $RequiredVersion is not installed."
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
        Stop-WithError "Module installation completed but $ModuleName $RequiredVersion is still unavailable."
    }

    Write-Success "$ModuleName $RequiredVersion installed."
    return [ordered]@{
        module = $ModuleName
        required_version = $RequiredVersion
        installed = $true
        installed_version = $installed.Version.ToString()
    }
}

function Get-ExchangeConnectCommand {
    param(
        [hashtable]$Configuration,
        [string]$AuthenticationMode
    )

    $environmentPart = if ([string]::IsNullOrWhiteSpace($Configuration.ExchangeEnvironmentName)) {
        ""
    }
    else {
        " -ExchangeEnvironmentName '$($Configuration.ExchangeEnvironmentName.Replace("'", "''"))'"
    }

    $delegatedPart = if ([string]::IsNullOrWhiteSpace($Configuration.DelegatedOrganization)) {
        ""
    }
    else {
        " -DelegatedOrganization '$($Configuration.DelegatedOrganization.Replace("'", "''"))'"
    }

    switch ($AuthenticationMode.ToLowerInvariant()) {
        "interactive" {
            $organizationPart = if ([string]::IsNullOrWhiteSpace($Configuration.ExchangeOrganization)) { "" } else { " -Organization '$($Configuration.ExchangeOrganization.Replace("'", "''"))'" }
            $upnHintPart = if ([string]::IsNullOrWhiteSpace($Configuration.UserPrincipalNameHint)) { "" } else { " -UserPrincipalName '$($Configuration.UserPrincipalNameHint.Replace("'", "''"))'" }
            return "Connect-ExchangeOnline -ShowBanner:`$false$environmentPart$organizationPart$delegatedPart$upnHintPart"
        }
        "devicecode" {
            $organizationPart = if ([string]::IsNullOrWhiteSpace($Configuration.ExchangeOrganization)) { "" } else { " -Organization '$($Configuration.ExchangeOrganization.Replace("'", "''"))'" }
            $upnHintPart = if ([string]::IsNullOrWhiteSpace($Configuration.UserPrincipalNameHint)) { "" } else { " -UserPrincipalName '$($Configuration.UserPrincipalNameHint.Replace("'", "''"))'" }
            return "Connect-ExchangeOnline -ShowBanner:`$false$environmentPart$organizationPart$delegatedPart$upnHintPart -Device"
        }
        "appcertificate" {
            if ([string]::IsNullOrWhiteSpace($Configuration.ApplicationId) -or [string]::IsNullOrWhiteSpace($Configuration.ExchangeOrganization)) {
                Stop-WithError "AppCertificate validation requires EXCHANGEADMIN_APP_ID and EXCHANGEADMIN_EXO_ORGANIZATION."
            }

            $certPart = if (-not [string]::IsNullOrWhiteSpace($Configuration.CertificateThumbprint)) {
                " -CertificateThumbprint '$($Configuration.CertificateThumbprint.Replace("'", "''"))'"
            }
            elseif (-not [string]::IsNullOrWhiteSpace($Configuration.CertificateSubjectName)) {
                " -CertificateSubjectName '$($Configuration.CertificateSubjectName.Replace("'", "''"))'"
            }
            else {
                Stop-WithError "AppCertificate validation requires EXCHANGEADMIN_CERT_THUMBPRINT or EXCHANGEADMIN_CERT_SUBJECT."
            }

            return "Connect-ExchangeOnline -ShowBanner:`$false$environmentPart$delegatedPart -AppId '$($Configuration.ApplicationId.Replace("'", "''"))' -Organization '$($Configuration.ExchangeOrganization.Replace("'", "''"))'$certPart"
        }
        "managedidentity" {
            if ([string]::IsNullOrWhiteSpace($Configuration.ExchangeOrganization)) {
                Stop-WithError "ManagedIdentity validation requires EXCHANGEADMIN_EXO_ORGANIZATION."
            }

            $accountIdPart = if ([string]::IsNullOrWhiteSpace($Configuration.ManagedIdentityAccountId)) { "" } else { " -ManagedIdentityAccountId '$($Configuration.ManagedIdentityAccountId.Replace("'", "''"))'" }
            return "Connect-ExchangeOnline -ShowBanner:`$false$environmentPart$delegatedPart -ManagedIdentity -Organization '$($Configuration.ExchangeOrganization.Replace("'", "''"))'$accountIdPart"
        }
        default {
            Stop-WithError "Unsupported authentication mode for tenant validation: $($Configuration.AuthenticationMode)"
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
            $tenantPart = if ([string]::IsNullOrWhiteSpace($Configuration.GraphTenantId)) { "" } else { " -TenantId '$($Configuration.GraphTenantId.Replace("'", "''"))'" }
            $scopeArguments = $scopes | ForEach-Object { "'$($_.Replace("'", "''"))'" }
            return "Connect-MgGraph -Scopes @($($scopeArguments -join ', '))$tenantPart -ContextScope Process -NoWelcome"
        }
        "devicecode" {
            $tenantPart = if ([string]::IsNullOrWhiteSpace($Configuration.GraphTenantId)) { "" } else { " -TenantId '$($Configuration.GraphTenantId.Replace("'", "''"))'" }
            $scopeArguments = $scopes | ForEach-Object { "'$($_.Replace("'", "''"))'" }
            return "Connect-MgGraph -Scopes @($($scopeArguments -join ', '))$tenantPart -ContextScope Process -UseDeviceCode -NoWelcome"
        }
        "appcertificate" {
            if ([string]::IsNullOrWhiteSpace($Configuration.ApplicationId) -or [string]::IsNullOrWhiteSpace($Configuration.GraphTenantId)) {
                Stop-WithError "AppCertificate validation requires EXCHANGEADMIN_APP_ID and EXCHANGEADMIN_GRAPH_TENANT_ID for Graph."
            }

            $certPart = if (-not [string]::IsNullOrWhiteSpace($Configuration.CertificateThumbprint)) {
                " -CertificateThumbprint '$($Configuration.CertificateThumbprint.Replace("'", "''"))'"
            }
            elseif (-not [string]::IsNullOrWhiteSpace($Configuration.CertificateSubjectName)) {
                " -CertificateSubjectName '$($Configuration.CertificateSubjectName.Replace("'", "''"))'"
            }
            else {
                Stop-WithError "AppCertificate validation requires EXCHANGEADMIN_CERT_THUMBPRINT or EXCHANGEADMIN_CERT_SUBJECT for Graph."
            }

            return "Connect-MgGraph -ClientId '$($Configuration.ApplicationId.Replace("'", "''"))' -TenantId '$($Configuration.GraphTenantId.Replace("'", "''"))'$certPart -ContextScope Process -NoWelcome"
        }
        "managedidentity" {
            $clientIdPart = if ([string]::IsNullOrWhiteSpace($Configuration.ManagedIdentityAccountId)) { "" } else { " -ClientId '$($Configuration.ManagedIdentityAccountId.Replace("'", "''"))'" }
            return "Connect-MgGraph -Identity$clientIdPart -ContextScope Process -NoWelcome"
        }
        default {
            Stop-WithError "Unsupported authentication mode for Graph validation: $($Configuration.AuthenticationMode)"
        }
    }
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
        [string[]]$RequiredParameters = @()
    )

    return Invoke-Probe -Name "Get-Command $CommandName" -ScriptBlock {
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

$resolvedConfigurationPath = if ([string]::IsNullOrWhiteSpace($ConfigurationPath)) { $null } else { Resolve-RepoPath -BaseDirectory $repositoryRoot -PathValue $ConfigurationPath }
$resolvedReportPath = Resolve-RepoPath -BaseDirectory $repositoryRoot -PathValue $ReportPath
$reportDirectory = Split-Path -Parent $resolvedReportPath
$bootstrapPolicyPath = Join-Path $repositoryRoot "src\ExchangeAdmin.Worker\Data\PowerShellModuleBootstrapPolicy.json"

if (-not (Test-Path $reportDirectory)) {
    New-Item -Path $reportDirectory -ItemType Directory -Force | Out-Null
}

$configuration = Load-ExchangeConfiguration -Path $resolvedConfigurationPath
$authResolution = Resolve-EffectiveAuthenticationMode -Configuration $configuration
$requiredModulesFromPolicy = Get-RequiredModulesFromPolicy -PolicyPath $bootstrapPolicyPath

$report = [ordered]@{
    generated_at_utc = (Get-Date).ToUniversalTime().ToString("O")
    configuration = [ordered]@{
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
    Write-Step "Ensuring approved PowerShell modules"
    $requiredModules = New-Object System.Collections.Generic.List[object]
    foreach ($moduleRequirement in $requiredModulesFromPolicy) {
        $moduleName = [string]$moduleRequirement.module
        if ($SkipGraph -and $moduleName.StartsWith("Microsoft.Graph.", [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }

        $requiredVersion = [string]$moduleRequirement.required_version
        if ([string]::IsNullOrWhiteSpace($requiredVersion)) {
            Stop-WithError "Bootstrap policy missing required version for $moduleName."
        }

        $requiredModules.Add((Ensure-RequiredModule -ModuleName $moduleName -RequiredVersion $requiredVersion -InstallIfMissing ([bool]$InstallApprovedModules)))
    }

    $report.prerequisites.required_modules = $requiredModules.ToArray()

    Write-Step "Connecting to Exchange Online"
    $exchangeConnectCommand = Get-ExchangeConnectCommand -Configuration $configuration -AuthenticationMode $authResolution.effective_mode
    Write-Info $exchangeConnectCommand
    Invoke-Expression $exchangeConnectCommand | Out-Null
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

    $report.compliance.probes.compliance_searches = Invoke-Probe -Name "Get-ComplianceSearch" -ScriptBlock {
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

    $report.compliance.probes.new_compliance_search = Invoke-CommandAvailabilityProbe -CommandName "New-ComplianceSearch" -RequiredParameters @("Name", "ExchangeLocation")
    $report.compliance.probes.start_compliance_search = Invoke-CommandAvailabilityProbe -CommandName "Start-ComplianceSearch" -RequiredParameters @("Identity")
    $report.compliance.probes.remove_compliance_search = Invoke-CommandAvailabilityProbe -CommandName "Remove-ComplianceSearch" -RequiredParameters @("Identity")
    $report.compliance.probes.new_compliance_search_action = Invoke-CommandAvailabilityProbe -CommandName "New-ComplianceSearchAction" -RequiredParameters @("SearchName", "Purge", "PurgeType")
    $report.compliance.probes.new_case_hold_policy = Invoke-CommandAvailabilityProbe -CommandName "New-CaseHoldPolicy" -RequiredParameters @("Name", "Case", "ExchangeLocation")
    $report.compliance.probes.new_case_hold_rule = Invoke-CommandAvailabilityProbe -CommandName "New-CaseHoldRule" -RequiredParameters @("Name", "Policy", "ContentMatchQuery")

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
        Write-Info $graphConnectCommand
        Invoke-Expression $graphConnectCommand | Out-Null
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

    $report | ConvertTo-Json -Depth 10 | Out-File -FilePath $resolvedReportPath -Encoding utf8
    Write-Info "Tenant validation report: $resolvedReportPath"
}
finally {
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    }
    catch {
    }

    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
    catch {
    }
}

if ($failedProbes.Count -gt 0) {
    Stop-WithError "Tenant validation completed with $($failedProbes.Count) failed probe(s). See report: $resolvedReportPath"
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " TENANT VALIDATION COMPLETED" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
