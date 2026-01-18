using System.Management.Automation;
using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

/// <summary>
/// Detects available Exchange Online cmdlets and their parameters.
/// </summary>
public class CapabilityDetector
{
    private readonly PowerShellEngine _engine;
    private CapabilityMapDto? _cachedCapabilities;

    /// <summary>
    /// List of cmdlets to detect.
    /// </summary>
    private static readonly string[] CmdletsToDetect = new[]
    {
        // Mailbox
        "Get-Mailbox",
        "Set-Mailbox",
        "Get-MailboxStatistics",
        "Get-MailboxPermission",
        "Add-MailboxPermission",
        "Remove-MailboxPermission",
        "Get-RecipientPermission",
        "Add-RecipientPermission",
        "Remove-RecipientPermission",
        "Enable-Mailbox",
        "Get-InboxRule",
        "Get-MailboxAutoReplyConfiguration",

        // Distribution Groups
        "Get-DistributionGroup",
        "Set-DistributionGroup",
        "Get-DistributionGroupMember",
        "Add-DistributionGroupMember",
        "Remove-DistributionGroupMember",

        // Dynamic Distribution Groups
        "Get-DynamicDistributionGroup",
        "Get-DynamicDistributionGroupMember",

        // Unified Groups (M365 Groups)
        "Get-UnifiedGroup",
        "Get-UnifiedGroupLinks",

        // Recipients
        "Get-Recipient",

        // Connection
        "Get-ConnectionInformation"
    };

    public CapabilityDetector(PowerShellEngine engine)
    {
        _engine = engine;
    }

    /// <summary>
    /// Gets cached capabilities or detects them.
    /// </summary>
    public CapabilityMapDto? CachedCapabilities => _cachedCapabilities;

    /// <summary>
    /// Detects available cmdlets and parameters.
    /// </summary>
    public async Task<CapabilityMapDto> DetectCapabilitiesAsync(
        bool forceRefresh = false,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        if (!forceRefresh && _cachedCapabilities != null)
        {
            return _cachedCapabilities;
        }

        onLog?.Invoke("Verbose", "Starting capability detection...");

        var capabilities = new CapabilityMapDto
        {
            DetectedAt = DateTime.UtcNow
        };

        // Build script to detect all cmdlets at once
        var script = BuildDetectionScript();

        var result = await _engine.ExecuteAsync(
            script,
            onVerbose: onLog,
            cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success || !result.Output.Any())
        {
            onLog?.Invoke("Warning", "Capability detection returned no results, using defaults");
            _cachedCapabilities = BuildDefaultCapabilities();
            return _cachedCapabilities;
        }

        // Parse results
        var detectedCmdlets = new Dictionary<string, CmdletCapabilityDto>();

        foreach (var output in result.Output)
        {
            if (output.BaseObject is System.Collections.Hashtable hash)
            {
                var name = hash["Name"]?.ToString();
                var isAvailable = hash["Available"] as bool? ?? false;
                var parameters = hash["Parameters"] as object[];
                var error = hash["Error"]?.ToString();

                if (!string.IsNullOrEmpty(name))
                {
                    detectedCmdlets[name] = new CmdletCapabilityDto
                    {
                        Name = name,
                        IsAvailable = isAvailable,
                        Parameters = parameters?.Select(p => p?.ToString() ?? "").Where(p => !string.IsNullOrEmpty(p)).ToList() ?? new List<string>(),
                        UnavailableReason = error
                    };

                    onLog?.Invoke("Verbose", $"  {name}: {(isAvailable ? "Available" : $"Not available - {error}")}");
                }
            }
        }

        capabilities.Cmdlets = detectedCmdlets;
        capabilities.Features = BuildFeatureCapabilities(detectedCmdlets);

        _cachedCapabilities = capabilities;

        onLog?.Invoke("Information", $"Capability detection complete: {detectedCmdlets.Count(c => c.Value.IsAvailable)} cmdlets available");

        return capabilities;
    }

    /// <summary>
    /// Builds the PowerShell script to detect cmdlets.
    /// </summary>
    private string BuildDetectionScript()
    {
        var cmdletList = string.Join("','", CmdletsToDetect);

        return $@"
$cmdlets = @('{cmdletList}')
$results = @()

foreach ($cmdletName in $cmdlets) {{
    try {{
        $cmd = Get-Command -Name $cmdletName -ErrorAction Stop
        $params = $cmd.Parameters.Keys | Where-Object {{ $_ -notlike '*Common*' }}
        $results += @{{
            Name = $cmdletName
            Available = $true
            Parameters = @($params)
            Error = $null
        }}
    }}
    catch {{
        $results += @{{
            Name = $cmdletName
            Available = $false
            Parameters = @()
            Error = $_.Exception.Message
        }}
    }}
}}

$results
";
    }

    /// <summary>
    /// Builds feature capabilities from detected cmdlets.
    /// </summary>
    private FeatureCapabilitiesDto BuildFeatureCapabilities(Dictionary<string, CmdletCapabilityDto> cmdlets)
    {
        bool IsAvailable(string name) => cmdlets.TryGetValue(name, out var c) && c.IsAvailable;
        bool HasParameter(string cmdlet, string param) =>
            cmdlets.TryGetValue(cmdlet, out var c) && c.IsAvailable && c.Parameters.Contains(param);

        return new FeatureCapabilitiesDto
        {
            // Mailbox features
            CanGetMailbox = IsAvailable("Get-Mailbox"),
            CanSetMailbox = IsAvailable("Set-Mailbox"),
            CanGetMailboxStatistics = IsAvailable("Get-MailboxStatistics"),
            CanGetMailboxPermission = IsAvailable("Get-MailboxPermission"),
            CanAddMailboxPermission = IsAvailable("Add-MailboxPermission"),
            CanRemoveMailboxPermission = IsAvailable("Remove-MailboxPermission"),
            CanGetRecipientPermission = IsAvailable("Get-RecipientPermission"),
            CanAddRecipientPermission = IsAvailable("Add-RecipientPermission"),
            CanRemoveRecipientPermission = IsAvailable("Remove-RecipientPermission"),

            // Mailbox feature toggles (check Set-Mailbox parameters)
            CanSetArchive = HasParameter("Set-Mailbox", "ArchiveDatabase") || IsAvailable("Enable-Mailbox"),
            CanSetLitigationHold = HasParameter("Set-Mailbox", "LitigationHoldEnabled"),
            CanSetAudit = HasParameter("Set-Mailbox", "AuditEnabled"),

            // Rules and auto-reply
            CanGetInboxRule = IsAvailable("Get-InboxRule"),
            CanGetMailboxAutoReplyConfiguration = IsAvailable("Get-MailboxAutoReplyConfiguration"),

            // Distribution groups
            CanGetDistributionGroup = IsAvailable("Get-DistributionGroup"),
            CanSetDistributionGroup = IsAvailable("Set-DistributionGroup"),
            CanGetDistributionGroupMember = IsAvailable("Get-DistributionGroupMember"),
            CanAddDistributionGroupMember = IsAvailable("Add-DistributionGroupMember"),
            CanRemoveDistributionGroupMember = IsAvailable("Remove-DistributionGroupMember"),

            // Dynamic distribution groups
            CanGetDynamicDistributionGroup = IsAvailable("Get-DynamicDistributionGroup"),
            CanGetDynamicDistributionGroupMember = IsAvailable("Get-DynamicDistributionGroupMember"),

            // Unified groups (M365 Groups)
            CanGetUnifiedGroup = IsAvailable("Get-UnifiedGroup"),
            CanGetUnifiedGroupLinks = IsAvailable("Get-UnifiedGroupLinks")
        };
    }

    /// <summary>
    /// Builds default capabilities when detection fails.
    /// </summary>
    private CapabilityMapDto BuildDefaultCapabilities()
    {
        var cmdlets = new Dictionary<string, CmdletCapabilityDto>();

        foreach (var cmdlet in CmdletsToDetect)
        {
            cmdlets[cmdlet] = new CmdletCapabilityDto
            {
                Name = cmdlet,
                IsAvailable = false,
                UnavailableReason = "Detection failed"
            };
        }

        return new CapabilityMapDto
        {
            DetectedAt = DateTime.UtcNow,
            Cmdlets = cmdlets,
            Features = new FeatureCapabilitiesDto()
        };
    }

    /// <summary>
    /// Clears cached capabilities.
    /// </summary>
    public void ClearCache()
    {
        _cachedCapabilities = null;
    }
}
