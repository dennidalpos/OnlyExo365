using System.Management.Automation;
using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

             
                                                                   
              
public class CapabilityDetector
{
    private readonly PowerShellEngine _engine;
    private CapabilityMapDto? _cachedCapabilities;

                 
                                  
                  
    private static readonly string[] CmdletsToDetect = new[]
    {
                  
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

                              
        "Get-DistributionGroup",
        "Set-DistributionGroup",
        "Get-DistributionGroupMember",
        "Add-DistributionGroupMember",
        "Remove-DistributionGroupMember",

                                      
        "Get-DynamicDistributionGroup",
        "Get-DynamicDistributionGroupMember",

                                       
        "Get-UnifiedGroup",
        "Get-UnifiedGroupLinks",

                     
        "Get-Recipient",

                     
        "Get-ConnectionInformation"
    };

    public CapabilityDetector(PowerShellEngine engine)
    {
        _engine = engine;
    }

                 
                                                 
                  
    public CapabilityMapDto? CachedCapabilities => _cachedCapabilities;

                 
                                                 
                  
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

                 
                                                          
                  
    private FeatureCapabilitiesDto BuildFeatureCapabilities(Dictionary<string, CmdletCapabilityDto> cmdlets)
    {
        bool IsAvailable(string name) => cmdlets.TryGetValue(name, out var c) && c.IsAvailable;
        bool HasParameter(string cmdlet, string param) =>
            cmdlets.TryGetValue(cmdlet, out var c) && c.IsAvailable && c.Parameters.Contains(param);

        return new FeatureCapabilitiesDto
        {
                               
            CanGetMailbox = IsAvailable("Get-Mailbox"),
            CanSetMailbox = IsAvailable("Set-Mailbox"),
            CanGetMailboxStatistics = IsAvailable("Get-MailboxStatistics"),
            CanGetMailboxPermission = IsAvailable("Get-MailboxPermission"),
            CanAddMailboxPermission = IsAvailable("Add-MailboxPermission"),
            CanRemoveMailboxPermission = IsAvailable("Remove-MailboxPermission"),
            CanGetRecipientPermission = IsAvailable("Get-RecipientPermission"),
            CanAddRecipientPermission = IsAvailable("Add-RecipientPermission"),
            CanRemoveRecipientPermission = IsAvailable("Remove-RecipientPermission"),

                                                                     
            CanSetArchive = HasParameter("Set-Mailbox", "ArchiveDatabase") || IsAvailable("Enable-Mailbox"),
            CanSetLitigationHold = HasParameter("Set-Mailbox", "LitigationHoldEnabled"),
            CanSetAudit = HasParameter("Set-Mailbox", "AuditEnabled"),

                                   
            CanGetInboxRule = IsAvailable("Get-InboxRule"),
            CanGetMailboxAutoReplyConfiguration = IsAvailable("Get-MailboxAutoReplyConfiguration"),

                                  
            CanGetDistributionGroup = IsAvailable("Get-DistributionGroup"),
            CanSetDistributionGroup = IsAvailable("Set-DistributionGroup"),
            CanGetDistributionGroupMember = IsAvailable("Get-DistributionGroupMember"),
            CanAddDistributionGroupMember = IsAvailable("Add-DistributionGroupMember"),
            CanRemoveDistributionGroupMember = IsAvailable("Remove-DistributionGroupMember"),

                                          
            CanGetDynamicDistributionGroup = IsAvailable("Get-DynamicDistributionGroup"),
            CanGetDynamicDistributionGroupMember = IsAvailable("Get-DynamicDistributionGroupMember"),

                                           
            CanGetUnifiedGroup = IsAvailable("Get-UnifiedGroup"),
            CanGetUnifiedGroupLinks = IsAvailable("Get-UnifiedGroupLinks")
        };
    }

                 
                                                         
                  
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

                 
                                   
                  
    public void ClearCache()
    {
        _cachedCapabilities = null;
    }
}
