#Requires -Version 7.6
[CmdletBinding()]
param([string]$RepositoryRoot, [string]$ScratchRoot, [string]$Scenario)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $RepositoryRoot 'scripts/internal/common.ps1')
. (Join-Path $RepositoryRoot 'build/helpers/cloud-routing.ps1')

function Assert-Contract([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}
function Import-Functions([string]$File) {
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $RepositoryRoot $File), [ref]$tokens, [ref]$errors)
    if ($errors.Count) { throw "Invalid script: $File" }
    foreach ($node in $ast.EndBlock.Statements) {
        if ($node -is [System.Management.Automation.Language.FunctionDefinitionAst]) {
            . ([scriptblock]::Create('function script:' + $node.Extent.Text.Substring('function '.Length)))
        }
    }
    return $ast
}
function New-Clone([string[]]$Files) {
    $root = Join-Path $ScratchRoot 'repo'
    foreach ($file in @('scripts/internal/common.ps1') + $Files) {
        $target = Join-Path $root $file
        New-Item -Path (Split-Path -Parent $target) -ItemType Directory -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $RepositoryRoot $file) -Destination $target
    }
    Copy-Item -LiteralPath (Join-Path $RepositoryRoot 'global.json') -Destination $root
    Set-Content -LiteralPath (Join-Path $root 'OnlyExo365.sln') -Value 'offline-fixture'
    return $root
}
function New-Marker([string]$Path, [string]$Value = 'retained') {
    New-Item -Path (Split-Path -Parent $Path) -ItemType Directory -Force | Out-Null
    Set-Content -LiteralPath $Path -Value $Value
}
function Invoke-Clean([string]$Root, [string[]]$Arguments = @()) {
    $script:cleanOutput = @(& pwsh -NoProfile -File (Join-Path $Root 'build/clean.ps1') @Arguments 2>&1)
    return $LASTEXITCODE
}

switch -Wildcard ($Scenario) {
    'clean-*' {
        $clone = New-Clone @('build/clean.ps1')
        $generated = Join-Path $clone 'artifacts/build/file.txt'
        New-Marker $generated
        switch ($Scenario) {
            'clean-preservation' {
                $markers = @('artifacts/security/scan.json','artifacts/test-results/result.trx','artifacts/exports/report.csv','artifacts/imports/input.csv','docs/example.log','docs/example.bak','unknown/bin/keep.txt')
                foreach ($file in $markers) { New-Marker (Join-Path $clone $file) }
                New-Marker (Join-Path $clone 'src/Example/Example.csproj')
                New-Marker (Join-Path $clone 'src/Example/obj/generated.txt')
                Assert-Contract ((Invoke-Clean $clone @('-SkipDotNetClean')) -eq 0) ($cleanOutput -join "`n")
                Assert-Contract (-not (Test-Path -LiteralPath $generated)) 'Generated output survived clean.'
                Assert-Contract (-not (Test-Path (Join-Path $clone 'src/Example/obj'))) 'Project obj survived clean.'
                foreach ($file in $markers) { Assert-Contract (Test-Path -LiteralPath (Join-Path $clone $file)) "User/evidence file deleted: $file" }
            }
            'clean-locked-file' {
                $handle = [IO.File]::Open($generated, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
                try { Assert-Contract ((Invoke-Clean $clone @('-SkipDotNetClean')) -ne 0) 'Locked deletion reported success.' }
                finally { $handle.Dispose() }
                Assert-Contract (Test-Path -LiteralPath $generated) 'Locked marker lost.'
            }
            'clean-unsafe-target' {
                $outside = Join-Path $ScratchRoot 'external/keep.txt'
                New-Marker $outside
                foreach ($target in @((Split-Path -Parent $outside), $clone, (Join-Path $clone 'artifacts'), (Join-Path $clone 'artifacts/exports/../..'))) {
                    Assert-Contract ((Invoke-Clean $clone @('-DryRun','-SkipDotNetClean','-ExportDirPath',$target)) -ne 0) "Unsafe target accepted: $target"
                    Assert-Contract (Test-Path -LiteralPath $generated) 'Invalid input caused partial cleanup.'
                }
            }
            'clean-reparse-point' {
                $external = Join-Path $ScratchRoot 'external'
                New-Marker (Join-Path $external 'keep.txt')
                New-Item -ItemType Junction -Path (Join-Path $clone 'artifacts/build/link') -Target $external | Out-Null
                Assert-Contract ((Invoke-Clean $clone @('-SkipDotNetClean')) -ne 0) 'Cleanup accepted a reparse point.'
                Assert-Contract (Test-Path (Join-Path $external 'keep.txt')) 'External data removed.'
                # Unlink the fixture junction before cleanup.
                [IO.Directory]::Delete((Join-Path $clone 'artifacts/build/link'))
            }
            'clean-tracked-output' {
                & git -C $clone init --quiet
                & git -C $clone add -f artifacts/build/file.txt
                if ($LASTEXITCODE) { throw 'Fixture git setup failed.' }
                Assert-Contract ((Invoke-Clean $clone @('-SkipDotNetClean')) -ne 0) 'Cleanup accepted tracked output.'
                Assert-Contract (Test-Path -LiteralPath $generated) 'Tracked output removed.'
            }
            'clean-native-failure' {
                function dotnet { $global:LASTEXITCODE = 37 }
                $failed = $false
                try { & (Join-Path $clone 'build/clean.ps1') -All } catch { $failed = $true }
                Assert-Contract $failed 'Native command failure was swallowed.'
            }
        }
    }
    'backup-*' {
        $null = Import-Functions 'build/run-smoke-tests.ps1'
        $null = Import-Functions 'build/helpers/run-smoke-tests.helpers.ps1'
        $data = Join-Path $ScratchRoot 'data'
        New-Marker (Join-Path $data 'marker.txt')
        $state = @{}
        if ($Scenario -eq 'backup-partial-failure') {
            $other = Join-Path $ScratchRoot 'other'
            New-Marker (Join-Path $other 'marker.txt')
            function Move-Item { param($LiteralPath, $Path, $Destination, [switch]$Force) if (($LiteralPath ?? $Path) -eq $other) { throw 'Injected second move failure' }; Microsoft.PowerShell.Management\Move-Item -LiteralPath ($LiteralPath ?? $Path) -Destination $Destination -Force:$Force }
            $backupFailed = $false
            try { $null = Backup-ProductDataRoots -Paths @($data,$other) -RecoveryState $state } catch { $backupFailed = $true }
            Assert-Contract $backupFailed 'Injected backup failure was swallowed.'
            Assert-Contract ($state.ContainsKey('ManifestPath') -and (Test-Path -LiteralPath $state.ManifestPath)) 'Recovery manifest missing after partial backup.'
            $count = Restore-ProductDataRoots -Entries @($state.Entries)
            Assert-Contract ($count -eq 2 -and (Test-Path (Join-Path $data 'marker.txt'))) 'Partial backup cannot recover all original roots.'
        }
        else {
            $backup = Backup-ProductDataRoots -Paths @($data) -RecoveryState $state
            if ($Scenario -eq 'backup-roundtrip') {
                New-Marker (Join-Path $data 'smoke-created.txt') 'new-data'
                New-Marker (Join-Path $backup.Entries[0].backup_path 'marker.txt') 'corrupted'
                $failed = $false
                try { $null = Restore-ProductDataRoots -Entries @($backup.Entries) -RecoveryState $state } catch { $failed = $true }
                Assert-Contract ($failed -and (Test-Path (Join-Path $data 'smoke-created.txt')) -and (Test-Path -LiteralPath $backup.Entries[0].backup_path)) 'Corrupted backup modified existing data or was discarded.'
                New-Marker (Join-Path $backup.Entries[0].backup_path 'marker.txt')
                Assert-Contract ((Restore-ProductDataRoots -Entries @($backup.Entries) -RecoveryState $state) -eq 1) 'Roundtrip restore count wrong.'
                Assert-Contract ((Get-Content (Join-Path $data 'marker.txt')).Trim() -eq 'retained') 'Restored data changed.'
                Assert-Contract (Test-Path (Join-Path $backup.Entries[0].created_data_path 'smoke-created.txt')) 'Recovery discarded data created during smoke execution.'
                Assert-Contract ((Restore-ProductDataRoots -Entries @($backup.Entries)) -eq 1) 'Recovery is not idempotent.'
            }
            elseif ($Scenario -eq 'backup-missing-entry') {
                [IO.Directory]::Move($backup.Entries[0].backup_path, (Join-Path $backup.BackupRoot 'hidden-backup'))
                $failed = $false
                try { $null = Restore-ProductDataRoots -Entries @($backup.Entries) } catch { $failed = $true }
                Assert-Contract $failed 'Missing backup and original were treated as restored.'
            }
            else {
                $installerSmokeRoot = $null
                $productDataBackup = $backup
                $ast = Import-Functions 'build/run-smoke-tests.ps1'
                function Restore-ProductDataRoots { param($Entries) throw 'Injected recovery failure' }
                $outer = @($ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.TryStatementAst] })[-1]
                $recoveryFailed = $false
                try { . ([scriptblock]::Create($outer.Finally.Extent.Text[1..($outer.Finally.Extent.Text.Length - 2)] -join '')) } catch { $recoveryFailed = $true }
                Assert-Contract $recoveryFailed 'Recovery failure was swallowed by finally.'
                Assert-Contract (Test-Path -LiteralPath $backup.Entries[0].backup_path) 'Failed recovery deleted only backup.'
            }
        }
        if ($state.ContainsKey('BackupRoot')) {
            # Fixture owns this root and unrecovered data.
            $tempRoot = [IO.Path]::GetFullPath($state.BackupRoot)
            Assert-Contract ($tempRoot.StartsWith([IO.Path]::GetTempPath(), [StringComparison]::OrdinalIgnoreCase) -and (Split-Path -Leaf $tempRoot).StartsWith('onlyexo365-localdata-backup-')) 'Unexpected fixture backup location.'
            [IO.Directory]::Delete($tempRoot, $true)
        }
    }
    'owned-process-cleanup' {
        $null = Import-Functions 'build/run-smoke-tests.ps1'
        $process = Start-Process -FilePath pwsh -ArgumentList @('-NoLogo','-NoProfile','-Command','Start-Sleep -Seconds 60') -WindowStyle Hidden -PassThru
        try {
            Stop-ProcessIfRunning -Process $process
            Assert-Contract $process.HasExited 'Session-owned test process survived cleanup.'
            Stop-ProcessIfRunning -Process $process
            Stop-ProcessIfRunning -Process $null
        }
        finally {
            if (-not $process.HasExited) { $process.Kill(); $process.WaitForExit() }
            $process.Dispose()
        }
    }
    'tenant-*' {
        $null = Import-Functions 'build/run-tenant-validation.ps1'
        function Get-EnvValue { param($Name) return $null }
        if ($Scenario -eq 'tenant-policy') {
            $modules = @(Get-RequiredModulesFromPolicy (Join-Path $RepositoryRoot 'src/OnlyExo365.Contracts/Data/PowerShellModuleBootstrapPolicy.json'))
            Assert-Contract ($modules.Count -eq 5) 'Canonical policy did not yield five exact module requirements.'
        }
        elseif ($Scenario -eq 'tenant-minimal-config') {
            $path = Join-Path $ScratchRoot 'minimal.json'
            Set-Content -LiteralPath $path -Value '{"ExchangeOnline":{"authenticationMode":"DeviceCode"}}'
            $config = Load-ExchangeConfiguration $path
            Assert-Contract ($config.AuthenticationMode -eq 'DeviceCode' -and $config.GraphScopes.Count -eq 4) 'Missing optional config properties broke defaults.'
        }
        elseif ($Scenario -eq 'tenant-app-only-coverage') {
            $clone = New-Clone @('build/run-tenant-validation.ps1','build/helpers/cloud-routing.ps1','src/OnlyExo365.Contracts/Data/CloudEnvironments.json','src/OnlyExo365.Contracts/Data/PowerShellModuleBootstrapPolicy.json')
            $scriptPath = Join-Path $clone 'build/run-tenant-validation.ps1'
            $source = Get-Content -LiteralPath $scriptPath -Raw
            $ast = Import-Functions 'build/run-tenant-validation.ps1'
            # Stub prerequisites; run actual report/probes.
            foreach ($node in $ast.EndBlock.Statements) {
                if ($node -isnot [Management.Automation.Language.FunctionDefinitionAst]) { continue }
                if ($node.Name -eq 'Ensure-RequiredModule') { $source = $source.Replace($node.Extent.Text, 'function Ensure-RequiredModule { param($ModuleName,$RequiredVersion,$InstallIfMissing) [pscustomobject]@{module=$ModuleName;version=$RequiredVersion} }') }
                if ($node.Name -eq 'Get-EnvValue') { $source = $source.Replace($node.Extent.Text, 'function Get-EnvValue { param($Name) return $null }') }
            }
            Set-Content -LiteralPath $scriptPath -Value $source
            function Import-Module { param($FullyQualifiedName,[switch]$Global,$ErrorAction) }
            function Connect-ExchangeOnline { param($AppId,$Organization,$CertificateThumbprint,$ShowBanner,$ExchangeEnvironmentName) }
            function Connect-IPPSSession { param($AppId,$Organization,$CertificateThumbprint,$ShowBanner,$AzureADAuthorizationEndpointUri) }
            function Disconnect-ExchangeOnline { param($Confirm,$ErrorAction) }
            function Get-OrganizationConfig { param($ErrorAction) [pscustomobject]@{Name='Offline fixture';PublicFoldersEnabled=$false} }
            function Get-AcceptedDomain { param($ErrorAction) @() }
            function Get-EXOMailbox { param($ResultSize,$Properties,$ErrorAction) @() }
            function Search-UnifiedAuditLog { param($StartDate,$EndDate,$ResultSize,$ErrorAction) @() }
            function Get-DkimSigningConfig { param($ErrorAction) @() }
            function Get-HostedContentFilterPolicy { param($ErrorAction) @() }
            function Get-AntiPhishPolicy { param($ErrorAction) @() }
            function Get-MalwareFilterPolicy { param($ErrorAction) @() }
            function Get-QuarantinePolicy { param($ErrorAction) @() }
            function Get-HostedOutboundSpamFilterPolicy { param($ErrorAction) @() }
            function New-CaseHoldPolicy { param($Name,$Case,$ExchangeLocation) throw 'Availability probes must never execute writes.' }
            function New-CaseHoldRule { param($Name,$Policy,$ContentMatchQuery) throw 'Availability probes must never execute writes.' }
            $configPath = Join-Path $ScratchRoot 'config.json'
            Set-Content -LiteralPath $configPath -Value '{"ExchangeOnline":{"authenticationMode":"AppCertificate","applicationId":"00000000-0000-0000-0000-000000000001","exchangeOrganization":"example.onmicrosoft.com","certificateThumbprint":"DUMMY"}}'
            $reportPath = Join-Path $ScratchRoot 'coverage.json'
            $failed = $false
            try { & $scriptPath -ConfigurationPath $configPath -ReportPath $reportPath -SkipGraph -InstallApprovedModules:$false } catch { $failed = $true }
            $report = Get-Content -LiteralPath $reportPath -Raw | ConvertFrom-Json
            Assert-Contract ($failed -and $report.exchange.connected -and $report.compliance.connected -and -not $report.compliance.search_session_connected) 'App-only coverage report has incorrect session state.'
            Assert-Contract ($report.summary.failed_probe_count -eq 6) 'Unsupported search session must fail all six search-dependent probes.'
            Assert-Contract (-not ($report.summary.PSObject.Properties.Name -contains 'fatal_error')) 'Offline tenant probe flow failed before producing coverage.'
        }
        elseif ($Scenario -eq 'tenant-failure-report') {
            $report = Join-Path $ScratchRoot 'failure.json'
            $output = @(& pwsh -NoProfile -File (Join-Path $RepositoryRoot 'build/run-tenant-validation.ps1') -ConfigurationPath (Join-Path $ScratchRoot 'missing.json') -ReportPath $report -InstallApprovedModules:$false 2>&1)
            Assert-Contract ($LASTEXITCODE -ne 0 -and (Test-Path -LiteralPath $report)) 'Preflight failure lost report or returned success.'
            $document = Get-Content -LiteralPath $report -Raw | ConvertFrom-Json
            Assert-Contract (-not [string]::IsNullOrWhiteSpace($document.summary.fatal_error) -and -not $document.exchange.connected) 'Preflight report claims a tenant connection.'
        }
        else {
            $config = Load-ExchangeConfiguration $null
            $config.ExchangeOrganization = 'example.onmicrosoft.com'
            $config.UserPrincipalNameHint = 'dummy@example.invalid'
            $config.GraphTenantId = '00000000-0000-0000-0000-000000000001'
            $exchange = Get-ExchangeConnectCommand $config Interactive
            Assert-Contract ($exchange.Arguments.Organization -eq $config.ExchangeOrganization -and $exchange.Arguments.UserPrincipalName -eq $config.UserPrincipalNameHint -and $exchange.Arguments.ExchangeEnvironmentName -eq 'O365Default') 'Optional Exchange arguments were discarded by dictionary conversion.'
            $graph = Get-GraphConnectCommand $config Interactive
            Assert-Contract ($graph.Arguments.TenantId -eq $config.GraphTenantId) 'Optional Graph tenant argument was discarded.'
            $command = Get-ComplianceConnectCommand $config Interactive -SearchOnly
            Assert-Contract ($command.CommandName -eq 'Connect-IPPSSession' -and $command.Arguments.EnableSearchOnlySession -and $command.Arguments.UserPrincipalName -eq $config.UserPrincipalNameHint) 'eDiscovery session arguments missing.'
            foreach ($mode in @('DeviceCode','AppCertificate','ManagedIdentity')) {
                $failed = $false
                try { $null = Get-ComplianceConnectCommand $config $mode -SearchOnly } catch { $failed = $true }
                Assert-Contract $failed "Unsupported eDiscovery authentication accepted: $mode"
            }
            $config.ExchangeEnvironmentName = 'O365China'
            $config.DelegatedOrganization = 'example.onmicrosoft.com'
            $command = Get-ComplianceConnectCommand $config Interactive -SearchOnly
            Assert-Contract ($command.Arguments.ConnectionUri -like 'https://*.outlook.cn/*' -and $command.Arguments.AzureADAuthorizationEndpointUri -eq 'https://login.chinacloudapi.cn/example.onmicrosoft.com') 'Compliance cloud or delegated tenant was discarded.'
        }
    }
    'signing-owned-certificate' {
        $ast = Import-Functions 'build/validate-artifact-signing.ps1'
        $script:removed = @()
        function Test-Path { param($Path, $LiteralPath) return $true }
        function Get-ChildItem { param($Path, $LiteralPath) @([pscustomobject]@{Thumbprint='OWNED'; PSPath='owned'},[pscustomobject]@{Thumbprint='OTHER'; PSPath='other'}) }
        function Remove-Item { param($Path, $LiteralPath, [switch]$Force, [switch]$DeleteKey, $ErrorAction) $script:removed += ($LiteralPath ?? $Path) }
        Remove-CertificateFromStore -StorePath 'dummy-store' -Thumbprint 'OWNED' -DeleteKey
        Assert-Contract (($removed -join ',') -eq 'owned') 'Cleanup selected another certificate.'
        Assert-Contract (-not ($ast.Extent.Text -match 'Remove-CertificatesBySubject')) 'Subject-wide certificate deletion remains.'
    }
    'installed-payload-signature' {
        $null = Import-Functions 'build/helpers/run-smoke-tests.helpers.ps1'
        $publish = Join-Path $ScratchRoot 'publish'
        $install = Join-Path $ScratchRoot 'install'
        foreach ($name in @('OnlyExo365.Shell.exe','OnlyExo365.Worker.exe')) {
            New-Marker (Join-Path $publish $name)
            New-Marker (Join-Path $install $name)
        }
        function Get-AuthenticodeSignature { param($LiteralPath) [pscustomobject]@{Status='Valid';SignerCertificate=[pscustomobject]@{Thumbprint='EXPECTED'}} }
        Assert-Contract (Assert-InstalledPayload $publish $install 'dummy-setup') 'Matching signed payload rejected.'
        function Get-AuthenticodeSignature { param($LiteralPath) [pscustomobject]@{Status='Valid';SignerCertificate=[pscustomobject]@{Thumbprint=$(if($LiteralPath -eq 'dummy-setup'){'EXPECTED'}else{'OTHER'})}} }
        $failed = $false
        try { $null = Assert-InstalledPayload $publish $install 'dummy-setup' } catch { $failed = $true }
        Assert-Contract $failed 'Installed payload with a different signer accepted.'
        New-Marker (Join-Path $install 'OnlyExo365.Worker.exe') 'altered'
        $failed = $false
        try { $null = Assert-InstalledPayload $publish $install 'dummy-setup' } catch { $failed = $true }
        Assert-Contract $failed 'Installer embedded a different payload.'
    }
    'catalog-*' {
        $null = Import-Functions 'scripts/agents/refresh-microsoft365-sku-catalog.ps1'
        $csv = 'Product_Display_Name,String_Id,GUID,Service_Plan_Name,Service_Plan_Id,Service_Plans_Included_Friendly_Names' + "`n" + '"Product, example",EXAMPLE,00000000-0000-0000-0000-000000000001,PLAN,00000000-0000-0000-0000-000000000002,"Plan, example"'
        if ($Scenario -eq 'catalog-single-sku') {
            $entries = @(Convert-CsvRowsToCatalogEntries @($csv | ConvertFrom-Csv))
            Assert-Contract ($entries.Count -eq 1 -and $entries[0].productName -eq 'Product, example') 'Catalog fixture mapping failed.'
            $clone = New-Clone @('scripts/agents/refresh-microsoft365-sku-catalog.ps1')
            function Invoke-WebRequest { param($Uri,$OutFile) Set-Content -LiteralPath $OutFile -Value $csv }
            $output = Join-Path $ScratchRoot 'catalog.json'
            & (Join-Path $clone 'scripts/agents/refresh-microsoft365-sku-catalog.ps1') -CsvDownloadUrl https://example.invalid/fixture.csv -OutputPath $output
            $document = Get-Content -LiteralPath $output -Raw | ConvertFrom-Json
            Assert-Contract ($document.entries -is [array] -and $document.entries.Count -eq 1) 'Single-SKU entries is not a JSON array.'
        }
        else {
            $failed = $false
            try { $null = Convert-CsvRowsToCatalogEntries @([pscustomobject]@{String_Id='BAD';GUID='not-a-guid';Product_Display_Name='';Service_Plan_Name='';Service_Plan_Id='';Service_Plans_Included_Friendly_Names=''}) } catch { $failed = $true }
            Assert-Contract $failed 'Invalid catalog accepted.'
        }
    }
    'reproducibility-evidence' {
        $clone = New-Clone @('build/verify-package-reproducibility.ps1','build/clean.ps1')
        New-Marker (Join-Path $clone 'artifacts/security/scan.json')
        @'
param($Configuration, [switch]$Clean, [switch]$LockedMode, $RuntimeIdentifier)
$root = Split-Path -Parent $PSScriptRoot
if ($Clean) { & (Join-Path $root 'build/clean.ps1') -SkipDotNetClean }
foreach ($directory in @('publish','packages')) {
    New-Item -Path (Join-Path $root "artifacts/$directory") -ItemType Directory -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $root "artifacts/$directory/file.txt") -Value 'deterministic'
}
'@ | Set-Content -LiteralPath (Join-Path $clone 'scripts/pack.ps1')
        $output = @(& pwsh -NoProfile -File (Join-Path $clone 'build/verify-package-reproducibility.ps1') 2>&1)
        Assert-Contract ($LASTEXITCODE -eq 0) ($output -join "`n")
        foreach ($path in @('artifacts/security/scan.json','artifacts/reproducibility/snapshots/run-1/publish/file.txt','artifacts/reproducibility/snapshots/run-2/publish/file.txt')) {
            Assert-Contract (Test-Path (Join-Path $clone $path)) "Reproducibility removed evidence: $path"
        }
    }
    default { throw "Unknown scenario: $Scenario" }
}
Write-Host "PASS: $Scenario"
