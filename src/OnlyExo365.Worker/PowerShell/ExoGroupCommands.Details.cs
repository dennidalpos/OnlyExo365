using System.Collections;
using System.Linq;
using OnlyExo365.Contracts.Dtos;

namespace OnlyExo365.Worker.PowerShell;

public partial class ExoGroupCommands
{
    public async Task<DistributionListDetailsDto> GetDistributionListDetailsAsync(
        GetDistributionListDetailsRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var normalizedGroupType = NormalizeRequestedGroupType(request.GroupTypeHint);
        return normalizedGroupType switch
        {
            GroupTypeMicrosoft365 => await GetUnifiedGroupDetailsAsync(request, onLog, cancellationToken),
            GroupTypeDynamic => await GetDynamicDistributionListDetailsAsync(request, onLog, cancellationToken),
            _ => await GetDistributionOrSecurityGroupDetailsAsync(request, onLog, cancellationToken)
        };
    }

    public async Task<GroupMembersPageDto> GetGroupMembersPageAsync(
        string identity,
        string groupType,
        int skip,
        int pageSize,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var normalizedGroupType = NormalizeRequestedGroupType(groupType);
        var escapedIdentity = EscapePowerShellString(identity);
        string script;
        if (normalizedGroupType == GroupTypeDynamic)
        {
            var capabilities = await _capabilityDetector.DetectCapabilitiesAsync(cancellationToken: cancellationToken);
            if (!capabilities.Features.CanGetDynamicDistributionGroup)
            {
            throw new InvalidOperationException("Dynamic member retrieval is not available: Get-DynamicDistributionGroup is not supported.");
            }

            var canUseDedicatedCmdlet = capabilities.Features.CanGetDynamicDistributionGroupMember;
            var canUseRecipientPreview = capabilities.Cmdlets.TryGetValue("Get-Recipient", out var getRecipientCapability) &&
                getRecipientCapability.IsAvailable;

            if (!canUseDedicatedCmdlet && !canUseRecipientPreview)
            {
            throw new InvalidOperationException("Dynamic member retrieval is not available: no supported membership path is available.");
            }

            script = BuildGetDynamicGroupMembersScript(escapedIdentity, skip, pageSize, canUseDedicatedCmdlet);
            onLog?.Invoke(
                "Verbose",
                canUseDedicatedCmdlet
                    ? $"Fetching dynamic group members via Get-DynamicDistributionGroupMember (skip={skip}, pageSize={pageSize})..."
                    : $"Fetching dynamic group members via Get-Recipient preview filter (skip={skip}, pageSize={pageSize})...");
        }
        else
        {
            script = normalizedGroupType switch
            {
                GroupTypeMicrosoft365 => BuildUnifiedGroupLinksScript(escapedIdentity, "Members", skip, pageSize),
                _ => $@"
$allMembers = Get-DistributionGroupMember -Identity '{escapedIdentity}' -ResultSize Unlimited

$totalCount = @($allMembers).Count
$pagedMembers = $allMembers | Select-Object -Skip {skip} -First {pageSize}

@{{
    TotalCount = $totalCount
    Members = @($pagedMembers | ForEach-Object {{
        @{{
            Identity = $_.Identity.ToString()
            Name = $_.DisplayName
            PrimarySmtpAddress = if ($_.PrimarySmtpAddress) {{ $_.PrimarySmtpAddress.ToString() }} else {{ $null }}
            RecipientType = if ($_.RecipientType) {{ $_.RecipientType.ToString() }} else {{ $null }}
        }}
    }})
}}
"
            };

            onLog?.Invoke("Verbose", $"Fetching group members (type={normalizedGroupType}, skip={skip}, pageSize={pageSize})...");
        }

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        return ParseGroupMembersPage(result, skip, pageSize);
    }

    private async Task<DistributionListDetailsDto> GetDistributionOrSecurityGroupDetailsAsync(
        GetDistributionListDetailsRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var escapedIdentity = EscapePowerShellString(request.Identity);
        var script = $@"
$group = Get-DistributionGroup -Identity '{escapedIdentity}'
$groupType = if ($group.RecipientTypeDetails -eq 'MailUniversalSecurityGroup') {{ '{GroupTypeMailSecurity}' }} else {{ '{GroupTypeDistribution}' }}

@{{
    Identity = $group.Identity.ToString()
    Guid = if ($group.Guid) {{ $group.Guid.ToString() }} else {{ $null }}
    DisplayName = $group.DisplayName
    PrimarySmtpAddress = if ($group.PrimarySmtpAddress) {{ $group.PrimarySmtpAddress.ToString() }} else {{ '' }}
    Alias = $group.Alias
    GroupType = $groupType
    RecipientType = if ($group.RecipientType) {{ $group.RecipientType.ToString() }} else {{ '' }}
    RecipientTypeDetails = if ($group.RecipientTypeDetails) {{ $group.RecipientTypeDetails.ToString() }} else {{ '' }}
    IsDynamic = $false
    EmailAddresses = @($group.EmailAddresses | ForEach-Object {{ $_.ToString() }})
    ManagedBy = @($group.ManagedBy | ForEach-Object {{ $_.ToString() }})
    AcceptMessagesOnlyFrom = @($group.AcceptMessagesOnlyFrom | ForEach-Object {{ $_.ToString() }})
    RejectMessagesFrom = @($group.RejectMessagesFrom | ForEach-Object {{ $_.ToString() }})
    RequireSenderAuthenticationEnabled = $group.RequireSenderAuthenticationEnabled
    HiddenFromAddressListsEnabled = $group.HiddenFromAddressListsEnabled
    MemberJoinRestriction = if ($group.MemberJoinRestriction) {{ $group.MemberJoinRestriction.ToString() }} else {{ $null }}
    MemberDepartRestriction = if ($group.MemberDepartRestriction) {{ $group.MemberDepartRestriction.ToString() }} else {{ $null }}
    WhenCreated = $group.WhenCreated
    WhenChanged = $group.WhenChanged
}}
";

        onLog?.Invoke("Verbose", $"Fetching group details for {request.Identity}...");
        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (!result.Output.Any() || result.Output.First().BaseObject is not Hashtable hash)
        {
            throw new InvalidOperationException($"Failed to get group details: {result.ErrorMessage}");
        }

        var details = CreateBaseGroupDetails(hash);
        if (request.IncludeMembers)
        {
            details.Members = await GetGroupMembersPageAsync(
                request.Identity,
                details.GroupType,
                0,
                request.MembersPageSize,
                onLog,
                cancellationToken);
        }

        return details;
    }

    private async Task<DistributionListDetailsDto> GetDynamicDistributionListDetailsAsync(
        GetDistributionListDetailsRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var escapedIdentity = EscapePowerShellString(request.Identity);
        var script = $@"
$group = Get-DynamicDistributionGroup -Identity '{escapedIdentity}'

@{{
    Identity = $group.Identity.ToString()
    Guid = if ($group.Guid) {{ $group.Guid.ToString() }} else {{ $null }}
    DisplayName = $group.DisplayName
    PrimarySmtpAddress = if ($group.PrimarySmtpAddress) {{ $group.PrimarySmtpAddress.ToString() }} else {{ '' }}
    Alias = $group.Alias
    GroupType = '{GroupTypeDynamic}'
    RecipientType = 'DynamicDistributionGroup'
    RecipientTypeDetails = 'DynamicDistributionGroup'
    IsDynamic = $true
    EmailAddresses = @($group.EmailAddresses | ForEach-Object {{ $_.ToString() }})
    ManagedBy = @($group.ManagedBy | ForEach-Object {{ $_.ToString() }})
    AcceptMessagesOnlyFrom = @($group.AcceptMessagesOnlyFrom | ForEach-Object {{ $_.ToString() }})
    AcceptMessagesOnlyFromSendersOrMembers = @($group.AcceptMessagesOnlyFromSendersOrMembers | ForEach-Object {{ $_.ToString() }})
    RejectMessagesFrom = @($group.RejectMessagesFrom | ForEach-Object {{ $_.ToString() }})
    RejectMessagesFromSendersOrMembers = @($group.RejectMessagesFromSendersOrMembers | ForEach-Object {{ $_.ToString() }})
    RequireSenderAuthenticationEnabled = $group.RequireSenderAuthenticationEnabled
    WhenCreated = $group.WhenCreated
    WhenChanged = $group.WhenChanged
}}
";

        onLog?.Invoke("Verbose", $"Fetching dynamic group details for {request.Identity}...");
        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (!result.Output.Any() || result.Output.First().BaseObject is not Hashtable hash)
        {
            throw new InvalidOperationException($"Failed to get dynamic group details: {result.ErrorMessage}");
        }

        var details = new DistributionListDetailsDto
        {
            Identity = hash["Identity"]?.ToString() ?? string.Empty,
            Guid = hash["Guid"]?.ToString(),
            DisplayName = hash["DisplayName"]?.ToString() ?? string.Empty,
            PrimarySmtpAddress = hash["PrimarySmtpAddress"]?.ToString() ?? string.Empty,
            Alias = hash["Alias"]?.ToString(),
            GroupType = GroupTypeDynamic,
            RecipientType = hash["RecipientType"]?.ToString() ?? "DynamicDistributionGroup",
            RecipientTypeDetails = hash["RecipientTypeDetails"]?.ToString() ?? "DynamicDistributionGroup",
            IsDynamic = true,
            EmailAddresses = ConvertToStringList(hash["EmailAddresses"]),
            ManagedBy = ConvertToStringList(hash["ManagedBy"]),
            AcceptMessagesOnlyFrom = MergeSenderLists(hash["AcceptMessagesOnlyFrom"], hash["AcceptMessagesOnlyFromSendersOrMembers"]),
            RejectMessagesFrom = MergeSenderLists(hash["RejectMessagesFrom"], hash["RejectMessagesFromSendersOrMembers"]),
            RequireSenderAuthenticationEnabled = hash["RequireSenderAuthenticationEnabled"] as bool? ?? false,
            WhenCreated = hash["WhenCreated"] as DateTime?,
            WhenChanged = hash["WhenChanged"] as DateTime?
        };

        if (request.IncludeMembers)
        {
            details.Members = await GetGroupMembersPageAsync(
                details.Identity,
                GroupTypeDynamic,
                0,
                request.MembersPageSize,
                onLog,
                cancellationToken);
        }

        return details;
    }

    private async Task<DistributionListDetailsDto> GetUnifiedGroupDetailsAsync(
        GetDistributionListDetailsRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var escapedIdentity = EscapePowerShellString(request.Identity);
        var script = $@"
$group = Get-UnifiedGroup -Identity '{escapedIdentity}'

@{{
    Identity = $group.Identity.ToString()
    Guid = if ($group.Guid) {{ $group.Guid.ToString() }} else {{ $null }}
    DisplayName = $group.DisplayName
    PrimarySmtpAddress = if ($group.PrimarySmtpAddress) {{ $group.PrimarySmtpAddress.ToString() }} else {{ '' }}
    Alias = $group.Alias
    GroupType = '{GroupTypeMicrosoft365}'
    RecipientType = if ($group.RecipientType) {{ $group.RecipientType.ToString() }} else {{ 'GroupMailbox' }}
    RecipientTypeDetails = if ($group.RecipientTypeDetails) {{ $group.RecipientTypeDetails.ToString() }} else {{ 'GroupMailbox' }}
    IsDynamic = $false
    EmailAddresses = @($group.EmailAddresses | ForEach-Object {{ $_.ToString() }})
    ManagedBy = @($group.ManagedBy | ForEach-Object {{ $_.ToString() }})
    AccessType = if ($group.AccessType) {{ $group.AccessType.ToString() }} else {{ $null }}
    Classification = $group.Classification
    Notes = $group.Notes
    HideFromAddressLists = $group.HiddenFromAddressListsEnabled
    HideFromExchangeClients = $group.HiddenFromExchangeClientsEnabled
    SubscriptionEnabled = $group.AutoSubscribeNewMembers
    WelcomeMessageEnabled = $group.WelcomeMessageEnabled
    WhenCreated = $group.WhenCreated
    WhenChanged = $group.WhenChanged
}}
";

        onLog?.Invoke("Verbose", $"Fetching Microsoft 365 group details for {request.Identity}...");
        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (!result.Output.Any() || result.Output.First().BaseObject is not Hashtable hash)
        {
            throw new InvalidOperationException($"Failed to get Microsoft 365 group details: {result.ErrorMessage}");
        }

        var details = CreateBaseGroupDetails(hash);
        if (request.IncludeMembers)
        {
            details.Members = await GetGroupMembersPageAsync(
                request.Identity,
                GroupTypeMicrosoft365,
                0,
                request.MembersPageSize,
                onLog,
                cancellationToken);
            details.Owners = await GetUnifiedGroupOwnersAsync(request.Identity, request.MembersPageSize, onLog, cancellationToken);
        }

        return details;
    }

    private async Task<GroupMembersPageDto> GetUnifiedGroupOwnersAsync(
        string identity,
        int pageSize,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var escapedIdentity = EscapePowerShellString(identity);
        var script = BuildUnifiedGroupLinksScript(escapedIdentity, "Owners", 0, pageSize);
        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        return ParseGroupMembersPage(result, 0, pageSize);
    }

    private static string BuildUnifiedGroupLinksScript(string escapedIdentity, string linkType, int skip, int pageSize)
    {
        return $@"
$allLinks = Get-UnifiedGroupLinks -Identity '{escapedIdentity}' -LinkType {linkType} -ResultSize Unlimited
$totalCount = @($allLinks).Count
$pagedLinks = $allLinks | Select-Object -Skip {skip} -First {pageSize}

@{{
    TotalCount = $totalCount
    Members = @($pagedLinks | ForEach-Object {{
        @{{
            Identity = $_.Identity.ToString()
            Name = if ($_.DisplayName) {{ $_.DisplayName }} else {{ $_.Name }}
            PrimarySmtpAddress = if ($_.PrimarySmtpAddress) {{ $_.PrimarySmtpAddress.ToString() }} else {{ $null }}
            RecipientType = if ($_.RecipientType) {{ $_.RecipientType.ToString() }} else {{ $null }}
        }}
    }})
}}
";
    }

    internal static string BuildGetDynamicGroupMembersScript(string escapedIdentity, int skip, int pageSize, bool useDedicatedCmdlet)
    {
        var membersScript = useDedicatedCmdlet
            ? $"$allMembers = Get-DynamicDistributionGroupMember -Identity '{escapedIdentity}' -ResultSize Unlimited"
            : $@"
$ddg = Get-DynamicDistributionGroup -Identity '{escapedIdentity}'
$recipientPreviewParams = @{{
    RecipientPreviewFilter = $ddg.RecipientFilter
    ResultSize = 'Unlimited'
}}

if (-not [string]::IsNullOrWhiteSpace($ddg.RecipientContainer)) {{
    $recipientPreviewParams['OrganizationalUnit'] = $ddg.RecipientContainer
}}

$allMembers = Get-Recipient @recipientPreviewParams";

        return $@"
{membersScript}

$totalCount = @($allMembers).Count
$pagedMembers = $allMembers | Select-Object -Skip {skip} -First {pageSize}

@{{
    TotalCount = $totalCount
    Members = @($pagedMembers | ForEach-Object {{
        @{{
            Identity = if ($_.Identity) {{ $_.Identity.ToString() }} elseif ($_.PrimarySmtpAddress) {{ $_.PrimarySmtpAddress.ToString() }} else {{ $_.Name }}
            Name = if ($_.DisplayName) {{ $_.DisplayName }} elseif ($_.Name) {{ $_.Name }} else {{ $_.Identity.ToString() }}
            PrimarySmtpAddress = if ($_.PrimarySmtpAddress) {{ $_.PrimarySmtpAddress.ToString() }} else {{ $null }}
            RecipientType = if ($_.RecipientType) {{ $_.RecipientType.ToString() }} else {{ $null }}
        }}
    }})
}}";
    }

    private static GroupMembersPageDto ParseGroupMembersPage(PowerShellResult result, int skip, int pageSize)
    {
        var page = new GroupMembersPageDto
        {
            Skip = skip,
            PageSize = pageSize
        };

        if (result.Success && result.Output.Any() && result.Output.First().BaseObject is Hashtable hash)
        {
            page.TotalCount = Convert.ToInt32(hash["TotalCount"] ?? 0);
            if (hash["Members"] is object[] members)
            {
                foreach (var memberHash in members.OfType<Hashtable>())
                {
                    page.Members.Add(ToGroupMember(memberHash));
                }
            }

            page.HasMore = (skip + page.Members.Count) < page.TotalCount;
        }

        return page;
    }

    private static DistributionListDetailsDto CreateBaseGroupDetails(Hashtable hash)
    {
        return new DistributionListDetailsDto
        {
            Identity = hash["Identity"]?.ToString() ?? string.Empty,
            Guid = hash["Guid"]?.ToString(),
            DisplayName = hash["DisplayName"]?.ToString() ?? string.Empty,
            PrimarySmtpAddress = hash["PrimarySmtpAddress"]?.ToString() ?? string.Empty,
            Alias = hash["Alias"]?.ToString(),
            GroupType = hash["GroupType"]?.ToString() ?? GroupTypeDistribution,
            RecipientType = hash["RecipientType"]?.ToString() ?? string.Empty,
            RecipientTypeDetails = hash["RecipientTypeDetails"]?.ToString() ?? string.Empty,
            IsDynamic = hash["IsDynamic"] as bool? ?? false,
            EmailAddresses = ConvertToStringList(hash["EmailAddresses"]),
            ManagedBy = ConvertToStringList(hash["ManagedBy"]),
            AcceptMessagesOnlyFrom = ConvertToStringList(hash["AcceptMessagesOnlyFrom"]),
            RejectMessagesFrom = ConvertToStringList(hash["RejectMessagesFrom"]),
            RequireSenderAuthenticationEnabled = hash["RequireSenderAuthenticationEnabled"] as bool? ?? false,
            HiddenFromAddressListsEnabled = hash["HiddenFromAddressListsEnabled"] as bool? ?? false,
            MemberJoinRestriction = hash["MemberJoinRestriction"]?.ToString(),
            MemberDepartRestriction = hash["MemberDepartRestriction"]?.ToString(),
            AccessType = hash["AccessType"]?.ToString(),
            Classification = hash["Classification"]?.ToString(),
            Notes = hash["Notes"]?.ToString(),
            HideFromAddressLists = hash["HideFromAddressLists"] as bool?,
            HideFromExchangeClients = hash["HideFromExchangeClients"] as bool?,
            SubscriptionEnabled = hash["SubscriptionEnabled"] as bool?,
            WelcomeMessageEnabled = hash["WelcomeMessageEnabled"] as bool?,
            WhenCreated = hash["WhenCreated"] as DateTime?,
            WhenChanged = hash["WhenChanged"] as DateTime?
        };
    }

    private static GroupMemberDto ToGroupMember(Hashtable hash)
    {
        return new GroupMemberDto
        {
            Identity = hash["Identity"]?.ToString() ?? string.Empty,
            Name = hash["Name"]?.ToString() ?? string.Empty,
            PrimarySmtpAddress = hash["PrimarySmtpAddress"]?.ToString(),
            RecipientType = hash["RecipientType"]?.ToString()
        };
    }
}

