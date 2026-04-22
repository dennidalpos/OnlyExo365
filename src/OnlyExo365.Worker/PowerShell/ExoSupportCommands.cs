using OnlyExo365.Contracts.Dtos;

namespace OnlyExo365.Worker.PowerShell;

internal sealed class ExoSupportCommands : ExoCommandModuleBase
{
    private static readonly HashSet<string> CompatibleExecutionPolicies = new(StringComparer.OrdinalIgnoreCase)
    {
        "Bypass",
        "Unrestricted",
        "RemoteSigned"
    };

    public ExoSupportCommands(PowerShellEngine engine)
        : base(engine)
    {
    }

    public async Task<PrerequisiteStatusDto> CheckPrerequisitesAsync(CancellationToken cancellationToken = default)
    {
        var exchangeModulePolicy = PowerShellModuleBootstrapPolicy.Resolve("ExchangeOnlineManagement")
            ?? throw new InvalidOperationException("Bootstrap policy missing ExchangeOnlineManagement entry.");
        var graphModulePolicy = PowerShellModuleBootstrapPolicy.Resolve("Microsoft.Graph")
            ?? throw new InvalidOperationException("Bootstrap policy missing Microsoft.Graph entry.");
        var escapedExchangeRequiredVersion = EscapePs(exchangeModulePolicy.RequiredVersion);
        var escapedGraphRequiredVersion = EscapePs(graphModulePolicy.RequiredVersion);
        var graphRequiredModulesLiteral = ToPsArrayLiteral(graphModulePolicy.GetRequiredModules());
        var script = @"
$psVersion = $PSVersionTable.PSVersion.ToString()
$isPwsh7 = $PSVersionTable.PSVersion.Major -ge 7
$currentUserExecutionPolicy = try { Get-ExecutionPolicy -Scope CurrentUser -ErrorAction Stop } catch { 'Unknown' }
$effectiveExecutionPolicy = try { Get-ExecutionPolicy -ErrorAction Stop } catch { 'Unknown' }
$exchangeModuleRequiredVersion = [Version]'" + escapedExchangeRequiredVersion + @"'
$graphModuleRequiredVersion = [Version]'" + escapedGraphRequiredVersion + @"'
$graphModuleNames = " + graphRequiredModulesLiteral + @"

$exoModule = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
$graphModules = foreach ($graphModuleName in $graphModuleNames) {
    Get-Module -ListAvailable -Name $graphModuleName | Sort-Object Version -Descending | Select-Object -First 1
}
$graphModuleMap = @{}
foreach ($graphModuleName in $graphModuleNames) {
    $module = $graphModules | Where-Object { $_.Name -eq $graphModuleName } | Select-Object -First 1
    $graphModuleMap[$graphModuleName] = if ($module) { $module.Version.ToString() } else { $null }
}
$graphModulesInstalled = @($graphModules).Count -eq $graphModuleNames.Count
$graphModulesApproved = $graphModulesInstalled -and (@($graphModules | Where-Object { $_.Version -ne $graphModuleRequiredVersion }).Count -eq 0)

[PSCustomObject]@{
    PowerShellVersion = $psVersion
    IsPowerShell7 = $isPwsh7
    CurrentUserExecutionPolicy = if ($currentUserExecutionPolicy) { $currentUserExecutionPolicy.ToString() } else { $null }
    EffectiveExecutionPolicy = if ($effectiveExecutionPolicy) { $effectiveExecutionPolicy.ToString() } else { $null }
    ExchangeModuleInstalled = ($null -ne $exoModule)
    ExchangeModuleVersion = if ($exoModule) { $exoModule.Version.ToString() } else { $null }
    ExchangeModuleRequiredVersion = $exchangeModuleRequiredVersion.ToString()
    IsExchangeModuleApproved = if ($exoModule) { $exoModule.Version -eq $exchangeModuleRequiredVersion } else { $false }
    GraphModuleInstalled = $graphModulesInstalled
    GraphModuleVersion = if ($graphModulesInstalled) { ($graphModuleMap.GetEnumerator() | Sort-Object Name | ForEach-Object { ""$($_.Name)=$($_.Value)"" }) -join '; ' } elseif (@($graphModules).Count -gt 0) { ($graphModuleMap.GetEnumerator() | Sort-Object Name | ForEach-Object { if ($_.Value) { ""$($_.Name)=$($_.Value)"" } else { ""$($_.Name)=missing"" } }) -join '; ' } else { $null }
    GraphModuleRequiredVersion = $graphModuleRequiredVersion.ToString()
    IsGraphModuleApproved = $graphModulesApproved
}";

        var results = await RunScriptAsync(script, cancellationToken);
        if (results.Count == 0)
        {
            return new PrerequisiteStatusDto();
        }

        var obj = results[0];
        var currentUserExecutionPolicy = GetNullableString(obj, "CurrentUserExecutionPolicy");
        var effectiveExecutionPolicy = GetNullableString(obj, "EffectiveExecutionPolicy");
        var isExecutionPolicyCompatible = IsExecutionPolicyCompatible(currentUserExecutionPolicy, effectiveExecutionPolicy);

        return new PrerequisiteStatusDto
        {
            PowerShellVersion = GetString(obj, "PowerShellVersion"),
            CurrentUserExecutionPolicy = currentUserExecutionPolicy,
            EffectiveExecutionPolicy = effectiveExecutionPolicy,
            IsExecutionPolicyCompatible = isExecutionPolicyCompatible,
            ManualInstructions = isExecutionPolicyCompatible ? null : BuildExecutionPolicyManualInstructions(),
            IsPowerShell7 = GetBool(obj, "IsPowerShell7"),
            ExchangeModuleInstalled = GetBool(obj, "ExchangeModuleInstalled"),
            ExchangeModuleVersion = GetNullableString(obj, "ExchangeModuleVersion"),
            ExchangeModuleRequiredVersion = GetNullableString(obj, "ExchangeModuleRequiredVersion"),
            IsExchangeModuleApproved = GetBool(obj, "IsExchangeModuleApproved"),
            GraphModuleInstalled = GetBool(obj, "GraphModuleInstalled"),
            GraphModuleVersion = GetNullableString(obj, "GraphModuleVersion"),
            GraphModuleRequiredVersion = GetNullableString(obj, "GraphModuleRequiredVersion"),
            IsGraphModuleApproved = GetBool(obj, "IsGraphModuleApproved")
        };
    }

    public async Task<InstallModuleResponse> InstallModuleAsync(InstallModuleRequest request, CancellationToken cancellationToken = default)
    {
        if (request == null)
        {
            throw new ArgumentNullException(nameof(request));
        }

        if (string.Equals(request.InstallTarget, "PowerShell7", StringComparison.OrdinalIgnoreCase))
        {
            return await InstallPowerShell7Async(request, cancellationToken);
        }

        var modulePolicy = PowerShellModuleBootstrapPolicy.Resolve(request.ModuleName);
        if (modulePolicy == null)
        {
            return new InstallModuleResponse
            {
                Success = false,
                Message = $"Unsupported module bootstrap request: {request.ModuleName}",
                ModuleName = request.ModuleName,
                ManualInstructions = "Manually install only modules approved by the repository, or update the worker embedded policy."
            };
        }

        var repositoryName = PowerShellModuleBootstrapPolicy.RepositoryName;
        var repositorySourceLocation = PowerShellModuleBootstrapPolicy.RepositorySourceLocation;
        var moduleName = modulePolicy.ModuleName;
        var installModuleNames = modulePolicy.GetRequiredModules();
        var safeModuleName = EscapePs(modulePolicy.ModuleName);
        var safeRequiredVersion = EscapePs(modulePolicy.RequiredVersion);
        var safeRepositoryName = EscapePs(repositoryName);
        var safeRepositorySourceLocation = EscapePs(repositorySourceLocation);
        var installModuleNamesLiteral = ToPsArrayLiteral(installModuleNames);
        var script = $@"
function Get-ExactInstalledModule {{
    param(
        [string]$Name,
        [Version]$RequiredVersion
    )

    return Get-Module -ListAvailable -Name $Name |
        Where-Object {{ $_.Version -eq $RequiredVersion }} |
        Sort-Object Version -Descending |
        Select-Object -First 1
}}

function Get-CurrentUserModulePath {{
    $documentsPath = [Environment]::GetFolderPath('MyDocuments')
    if ([string]::IsNullOrWhiteSpace($documentsPath)) {{
        throw 'Unable to resolve the current user Documents folder.'
    }}

    return Join-Path (Join-Path $documentsPath 'PowerShell') 'Modules'
}}

function Install-ModuleFromGalleryPackage {{
    param(
        [string]$Name,
        [string]$RequiredVersion,
        [string]$RepositorySourceLocation
    )

    $downloadUri = ""$($RepositorySourceLocation.TrimEnd('/'))/package/$Name/$RequiredVersion""
    $tempRoot = Join-Path ([IO.Path]::GetTempPath()) ([Guid]::NewGuid().ToString('N'))
    $packagePath = Join-Path $tempRoot ""$Name.$RequiredVersion.nupkg""
    $expandedPath = Join-Path $tempRoot 'expanded'
    $targetDirectory = Join-Path (Join-Path (Get-CurrentUserModulePath) $Name) $RequiredVersion

    try {{
        Add-Type -AssemblyName System.Net.Http
        Add-Type -AssemblyName System.IO.Compression.FileSystem

        New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

        $handler = [System.Net.Http.HttpClientHandler]::new()
        $handler.AllowAutoRedirect = $true
        $client = [System.Net.Http.HttpClient]::new($handler)
        try {{
            $client.DefaultRequestHeaders.UserAgent.ParseAdd('OnlyExo365/OnlyExo365')
            $response = $client.GetAsync($downloadUri).GetAwaiter().GetResult()
            $response.EnsureSuccessStatusCode()

            $packageBytes = $response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult()
            if (-not $packageBytes -or $packageBytes.Length -le 0) {{
                $finalUri = if ($response.RequestMessage -and $response.RequestMessage.RequestUri) {{ $response.RequestMessage.RequestUri.ToString() }} else {{ $downloadUri }}
                throw ""Downloaded package for $Name $RequiredVersion is empty. Final URI: $finalUri""
            }}

            [System.IO.File]::WriteAllBytes($packagePath, $packageBytes)
        }}
        finally {{
            $client.Dispose()
            $handler.Dispose()
        }}

        [System.IO.Compression.ZipFile]::ExtractToDirectory($packagePath, $expandedPath)

        $rootManifestPath = Join-Path $expandedPath ""$Name.psd1""
        $manifest = if (Test-Path $rootManifestPath) {{
            Get-Item $rootManifestPath -ErrorAction Stop
        }}
        else {{
            Get-ChildItem -Path $expandedPath -Recurse -File -Include '*.psd1' -ErrorAction Stop |
                Where-Object {{ $_.Name -eq ""$Name.psd1"" }} |
                Sort-Object FullName |
                Select-Object -First 1
        }}

        if (-not $manifest) {{
            $manifest = Get-ChildItem -Path $expandedPath -Recurse -File -Include '*.psd1' -ErrorAction Stop |
                Sort-Object FullName |
                Select-Object -First 1
        }}

        if (-not $manifest) {{
            $topLevelEntries = Get-ChildItem -Path $expandedPath -Force -ErrorAction SilentlyContinue |
                Select-Object -ExpandProperty Name
            $topLevelSummary = if ($topLevelEntries) {{ [string]::Join(', ', $topLevelEntries) }} else {{ '<empty>' }}
            throw ""Downloaded package for $Name $RequiredVersion does not contain a module manifest. Top-level entries: $topLevelSummary""
        }}

        $moduleContentRoot = $manifest.Directory.FullName
        if (Test-Path $targetDirectory) {{
            Remove-Item -Path $targetDirectory -Recurse -Force -ErrorAction Stop
        }}

        New-Item -ItemType Directory -Path $targetDirectory -Force | Out-Null
        Get-ChildItem -Path $moduleContentRoot -Force -ErrorAction Stop | ForEach-Object {{
            Copy-Item -Path $_.FullName -Destination $targetDirectory -Recurse -Force -ErrorAction Stop
        }}
    }}
    finally {{
        if (Test-Path $tempRoot) {{
            Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }}
    }}
}}

try {{
    $ConfirmPreference = 'None'
    $moduleName = '{safeModuleName}'
    $moduleNames = {installModuleNamesLiteral}
    $requiredVersion = '{safeRequiredVersion}'
    $requiredVersionObject = [Version]$requiredVersion
    $repositoryName = '{safeRepositoryName}'
    $repositorySourceLocation = '{safeRepositorySourceLocation}'

    $installPsResource = Get-Command Install-PSResource -ErrorAction SilentlyContinue
    $getPsResourceRepository = Get-Command Get-PSResourceRepository -ErrorAction SilentlyContinue
    if ($installPsResource -and $getPsResourceRepository) {{
        $repo = Get-PSResourceRepository -Name $repositoryName -ErrorAction SilentlyContinue
        $repoUri = if ($repo -and $repo.Uri) {{ $repo.Uri.ToString() }} else {{ $null }}
        if ($repoUri -and $repoUri.TrimEnd('/') -ne $repositorySourceLocation.TrimEnd('/')) {{
            throw ""Repository $repositoryName source mismatch: $repoUri""
        }}
    }}

    $alreadyInstalled = foreach ($requestedModule in $moduleNames) {{
        Get-ExactInstalledModule -Name $requestedModule -RequiredVersion $requiredVersionObject
    }}

    if (@($alreadyInstalled).Count -eq $moduleNames.Count) {{
        [PSCustomObject]@{{
            Success = $true
            Message = ""$($moduleNames -join ', ') approved version $requiredVersion already installed""
            InstalledVersion = $requiredVersion
        }}
        return
    }}

    foreach ($requestedModule in $moduleNames) {{
        $installedModule = Get-ExactInstalledModule -Name $requestedModule -RequiredVersion $requiredVersionObject
        if ($installedModule) {{
            continue
        }}

        if ($installPsResource -and $getPsResourceRepository) {{
        $installPsResourceArguments = @{{
                Name = $requestedModule
            Version = $requiredVersion
            Repository = $repositoryName
            Scope = 'CurrentUser'
            TrustRepository = $true
            Quiet = $true
            ErrorAction = 'Stop'
        }}

        if ($installPsResource.Parameters.ContainsKey('AcceptLicense')) {{
            $installPsResourceArguments['AcceptLicense'] = $true
        }}

            Install-PSResource @installPsResourceArguments
        }}
        else {{
            Install-ModuleFromGalleryPackage -Name $requestedModule -RequiredVersion $requiredVersion -RepositorySourceLocation $repositorySourceLocation
        }}
    }}

    $installed = foreach ($requestedModule in $moduleNames) {{
        Get-ExactInstalledModule -Name $requestedModule -RequiredVersion $requiredVersionObject
    }}
    [PSCustomObject]@{{
        Success = (@($installed).Count -eq $moduleNames.Count)
        Message = if (@($installed).Count -eq $moduleNames.Count) {{ ""$($moduleNames -join ', ') approved version $requiredVersion installed successfully"" }} else {{ ""Module install completed but one or more required modules were not found in PSModulePath: $($moduleNames -join ', ')"" }}
        InstalledVersion = if (@($installed).Count -eq $moduleNames.Count) {{ $requiredVersion }} else {{ $null }}
    }}
}} catch {{
    [PSCustomObject]@{{
        Success = $false
        Message = ""Failed to install {safeModuleName} approved version {safeRequiredVersion}: $($_.Exception.Message)""
        InstalledVersion = $null
    }}
}}";

        var result = await Engine.ExecuteAsync(script, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (result.Output.Count == 0)
        {
            return new InstallModuleResponse
            {
                Success = false,
                Message = result.ErrorMessage ?? $"No output from Install-Module {moduleName}",
                ModuleName = moduleName
            };
        }

        var obj = result.Output.Last();
        var success = GetBool(obj, "Success");
        var message = GetString(obj, "Message");
        if (!success && string.IsNullOrWhiteSpace(message))
        {
            message = result.ErrorMessage ?? $"Failed to install {moduleName}";
        }

        return new InstallModuleResponse
        {
            Success = success,
            Message = message,
            ModuleName = moduleName,
            InstalledVersion = GetNullableString(obj, "InstalledVersion"),
            ManualInstructions = success ? null : BuildModuleInstallManualInstructions(modulePolicy, repositoryName)
        };
    }

    private async Task<InstallModuleResponse> InstallPowerShell7Async(InstallModuleRequest request, CancellationToken cancellationToken)
    {
        var packageId = string.IsNullOrWhiteSpace(request.PackageId)
            ? "Microsoft.PowerShell"
            : request.PackageId.Trim();

        var escapedPackageId = EscapePs(packageId);
        var script = $@"
function Get-PowerShell7Installation {{
    $pwshCommand = Get-Command pwsh -ErrorAction SilentlyContinue
    if ($pwshCommand) {{
        return [PSCustomObject]@{{
            Found = $true
            Version = try {{ & $pwshCommand.Source -NoLogo -NoProfile -Command ""$PSVersionTable.PSVersion.ToString()"" }} catch {{ $null }}
            Location = $pwshCommand.Source
        }}
    }}

    $candidatePaths = @(
        Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'
        Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\pwsh.exe'
        Join-Path $env:LOCALAPPDATA 'Programs\PowerShell\7\pwsh.exe'
    ) | Where-Object {{ $_ -and (Test-Path $_) }} | Select-Object -Unique

    foreach ($candidate in $candidatePaths) {{
        return [PSCustomObject]@{{
            Found = $true
            Version = try {{ & $candidate -NoLogo -NoProfile -Command ""$PSVersionTable.PSVersion.ToString()"" }} catch {{ $null }}
            Location = $candidate
        }}
    }}

    return [PSCustomObject]@{{
        Found = $false
        Version = $null
        Location = $null
    }}
}}

try {{
    $existing = Get-PowerShell7Installation
    if ($existing.Found) {{
        [PSCustomObject]@{{
            Success = $true
            Message = 'PowerShell 7 already installed'
            InstalledVersion = $existing.Version
            ManualInstructions = $null
        }}
        return
    }}

    $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
    if (-not $winget) {{
        [PSCustomObject]@{{
            Success = $false
            Message = 'winget is not available on this system'
            InstalledVersion = $null
            ManualInstructions = 'Manually install PowerShell 7 from https://github.com/PowerShell/PowerShell/releases, or enable winget and rerun: winget install --id Microsoft.PowerShell --exact --source winget'
        }}
        return
    }}

    $arguments = @(
        'install',
        '--id', '{escapedPackageId}',
        '--exact',
        '--source', 'winget',
        '--accept-package-agreements',
        '--accept-source-agreements',
        '--silent'
    )

    $process = Start-Process -FilePath $winget.Source -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden
    if ($process.ExitCode -ne 0) {{
        [PSCustomObject]@{{
            Success = $false
            Message = ""winget exited with code $($process.ExitCode)""
            InstalledVersion = $null
            ManualInstructions = 'Run manually: winget install --id Microsoft.PowerShell --exact --source winget, or download the package from https://github.com/PowerShell/PowerShell/releases'
        }}
        return
    }}

    $installed = Get-PowerShell7Installation
    [PSCustomObject]@{{
        Success = $installed.Found
        Message = if ($installed.Found) {{ 'PowerShell 7 installed via winget' }} else {{ 'winget completed but PowerShell 7 was not detected afterwards' }}
        InstalledVersion = $installed.Version
        ManualInstructions = if ($installed.Found) {{ $null }} else {{ 'Verify the installation in Apps & Features, or run manually: winget install --id Microsoft.PowerShell --exact --source winget' }}
    }}
}} catch {{
    [PSCustomObject]@{{
        Success = $false
        Message = ""Failed to install PowerShell 7: $($_.Exception.Message)""
        InstalledVersion = $null
        ManualInstructions = 'Run manually: winget install --id Microsoft.PowerShell --exact --source winget, or download the package from https://github.com/PowerShell/PowerShell/releases'
    }}
}}";

        var result = await Engine.ExecuteAsync(script, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (result.Output.Count == 0)
        {
            return new InstallModuleResponse
            {
                Success = false,
                Message = result.ErrorMessage ?? "No output from PowerShell 7 installer",
                ModuleName = request.ModuleName,
                ManualInstructions = BuildPowerShell7ManualInstructions()
            };
        }

        var obj = result.Output.Last();
        var success = GetBool(obj, "Success");
        var message = GetString(obj, "Message");

        return new InstallModuleResponse
        {
            Success = success,
            Message = string.IsNullOrWhiteSpace(message)
                ? (success ? "PowerShell 7 installed" : "PowerShell 7 installation failed")
                : message,
            ModuleName = request.ModuleName,
            InstalledVersion = GetNullableString(obj, "InstalledVersion"),
            ManualInstructions = success ? null : GetNullableString(obj, "ManualInstructions") ?? BuildPowerShell7ManualInstructions()
        };
    }

    private static bool IsExecutionPolicyCompatible(string? currentUserExecutionPolicy, string? effectiveExecutionPolicy)
    {
        return IsCompatibleExecutionPolicyValue(currentUserExecutionPolicy) ||
               IsCompatibleExecutionPolicyValue(effectiveExecutionPolicy);
    }

    private static bool IsCompatibleExecutionPolicyValue(string? value)
    {
        return !string.IsNullOrWhiteSpace(value) && CompatibleExecutionPolicies.Contains(value.Trim());
    }

    private static string BuildExecutionPolicyManualInstructions()
    {
        return "Open PowerShell 7 as the current user and run: Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned. In managed environments, check corporate policy or GPO requirements first.";
    }

    private static string BuildModuleInstallManualInstructions(PowerShellModuleDefinition modulePolicy, string repositoryName)
    {
        return modulePolicy.BuildManualInstructions(repositoryName);
    }

    private static string BuildPowerShell7ManualInstructions()
    {
        return "Run `winget install --id Microsoft.PowerShell --exact --source winget`, or download the package from https://github.com/PowerShell/PowerShell/releases";
    }

    private static string ToPsArrayLiteral(IReadOnlyList<string> values)
    {
        if (values.Count == 0)
        {
            return "@()";
        }

        return $"@({string.Join(", ", values.Select(value => $"'{EscapePs(value)}'"))})";
    }
}

