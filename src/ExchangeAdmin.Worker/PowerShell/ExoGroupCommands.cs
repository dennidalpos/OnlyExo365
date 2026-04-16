using System.Collections;
using System.Linq;
using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

public partial class ExoGroupCommands
{
    private const string GroupTypeDistribution = "Distribution";
    private const string GroupTypeMailSecurity = "MailSecurity";
    private const string GroupTypeDynamic = "Dynamic";
    private const string GroupTypeMicrosoft365 = "Microsoft365";

    private readonly PowerShellEngine _engine;
    private readonly CapabilityDetector _capabilityDetector;

    public ExoGroupCommands(PowerShellEngine engine, CapabilityDetector capabilityDetector)
    {
        _engine = engine;
        _capabilityDetector = capabilityDetector;
    }

    public async Task ModifyGroupMemberAsync(
        ModifyGroupMemberRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var normalizedGroupType = NormalizeRequestedGroupType(request.GroupType);
        if (normalizedGroupType == GroupTypeDynamic)
        {
            throw new InvalidOperationException("Dynamic group members cannot be modified manually.");
        }

        var escapedIdentity = EscapePowerShellString(request.Identity);
        var escapedMember = EscapePowerShellString(request.Member);
        var actionVerb = request.Action == GroupMemberAction.Add ? "Adding" : "Removing";

        string script = normalizedGroupType switch
        {
            GroupTypeMicrosoft365 when request.Action == GroupMemberAction.Add =>
                $"Add-UnifiedGroupLinks -Identity '{escapedIdentity}' -LinkType Members -Links '{escapedMember}'",
            GroupTypeMicrosoft365 =>
                $"Remove-UnifiedGroupLinks -Identity '{escapedIdentity}' -LinkType Members -Links '{escapedMember}' -Confirm:$false",
            _ when request.Action == GroupMemberAction.Add =>
                $"Add-DistributionGroupMember -Identity '{escapedIdentity}' -Member '{escapedMember}' -Confirm:$false",
            _ =>
                $"Remove-DistributionGroupMember -Identity '{escapedIdentity}' -Member '{escapedMember}' -Confirm:$false"
        };

        onLog?.Invoke("Information", $"{actionVerb} {request.Member} {(request.Action == GroupMemberAction.Add ? "to" : "from")} {request.Identity}...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to {request.Action} member: {result.ErrorMessage}");
        }

        onLog?.Invoke("Information", $"Successfully {(request.Action == GroupMemberAction.Add ? "added" : "removed")} member");
    }

    public async Task CreateDistributionListAsync(
        CreateDistributionListRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(request.DisplayName) ||
            string.IsNullOrWhiteSpace(request.Alias) ||
            string.IsNullOrWhiteSpace(request.PrimarySmtpAddress))
        {
            throw new InvalidOperationException("DisplayName, Alias and PrimarySmtpAddress are required to create a distribution list.");
        }

        var escapedDisplayName = EscapePowerShellString(request.DisplayName);
        var escapedAlias = EscapePowerShellString(request.Alias);
        var escapedPrimarySmtpAddress = EscapePowerShellString(request.PrimarySmtpAddress);
        var script =
            $"New-DistributionGroup -Name '{escapedDisplayName}' -DisplayName '{escapedDisplayName}' -Alias '{escapedAlias}' -PrimarySmtpAddress '{escapedPrimarySmtpAddress}'";

        onLog?.Invoke("Information", $"Creating distribution list {request.PrimarySmtpAddress}...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create distribution list: {result.ErrorMessage}");
        }

        onLog?.Invoke("Information", "Distribution list created successfully");
    }

    public async Task<PreviewDynamicGroupMembersResponse> PreviewDynamicGroupMembersAsync(
        PreviewDynamicGroupMembersRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var capabilities = await _capabilityDetector.DetectCapabilitiesAsync(cancellationToken: cancellationToken);
        if (!capabilities.Features.CanGetDynamicDistributionGroup)
        {
            throw new InvalidOperationException("Dynamic member preview is not available: Get-DynamicDistributionGroup is not supported.");
        }

        var canUseDedicatedCmdlet = capabilities.Features.CanGetDynamicDistributionGroupMember;
        var canGetRecipient = capabilities.Cmdlets.TryGetValue("Get-Recipient", out var getRecipientCapability) &&
            getRecipientCapability.IsAvailable;
        if (!canUseDedicatedCmdlet && !canGetRecipient)
        {
            throw new InvalidOperationException("Dynamic member preview is not available: no supported membership path is available.");
        }

        var escapedIdentity = EscapePowerShellString(request.Identity);
        var script = BuildGetDynamicGroupMembersScript(
            escapedIdentity,
            skip: 0,
            pageSize: request.MaxResults,
            useDedicatedCmdlet: canUseDedicatedCmdlet);

        onLog?.Invoke(
            "Warning",
            canUseDedicatedCmdlet
                ? "Previewing dynamic group members via Get-DynamicDistributionGroupMember..."
                : "Previewing dynamic group members via Get-Recipient preview filter (may be slow for large groups)...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        var response = new PreviewDynamicGroupMembersResponse
        {
            Identity = request.Identity
        };

        if (result.Success && result.Output.Any() && result.Output.First().BaseObject is Hashtable hash)
        {
            response.TotalCount = Convert.ToInt32(hash["TotalCount"] ?? 0);
            response.IsLimited = response.TotalCount > request.MaxResults;
            if (response.IsLimited)
            {
                response.Warning = $"Preview limited to {request.MaxResults} of {response.TotalCount} members";
            }

            if (hash["Members"] is object[] members)
            {
                foreach (var memberHash in members.OfType<Hashtable>())
                {
                    response.Members.Add(ToGroupMember(memberHash));
                }
            }
        }

        onLog?.Invoke("Information", $"Preview complete: {response.Members.Count} members shown (total: {response.TotalCount})");
        return response;
    }

    public async Task SetDistributionListSettingsAsync(
        SetDistributionListSettingsRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var normalizedGroupType = NormalizeRequestedGroupType(request.GroupType);
        if (normalizedGroupType == GroupTypeMicrosoft365)
        {
            throw new InvalidOperationException("Advanced delivery settings are not supported for Microsoft 365 Groups in this module.");
        }

        var escapedIdentity = EscapePowerShellString(request.Identity);
        var setParams = new List<string>();
        var isDynamicGroup = normalizedGroupType == GroupTypeDynamic;
        var cmdletName = isDynamicGroup ? "Set-DynamicDistributionGroup" : "Set-DistributionGroup";

        var capabilities = await _capabilityDetector.DetectCapabilitiesAsync(cancellationToken: cancellationToken);
        var cmdletParameters = capabilities.Cmdlets.TryGetValue(cmdletName, out var cmdletCapability)
            ? cmdletCapability.Parameters
            : new List<string>();

        bool SupportsParam(string paramName) => cmdletParameters.Contains(paramName, StringComparer.OrdinalIgnoreCase);

        if (request.RequireSenderAuthenticationEnabled.HasValue)
        {
            if (isDynamicGroup
                ? capabilities.Features.CanSetDynamicDistributionGroupRequireSenderAuthentication
                : capabilities.Features.CanSetDistributionGroupRequireSenderAuthentication)
            {
                setParams.Add($"-RequireSenderAuthenticationEnabled ${request.RequireSenderAuthenticationEnabled.Value.ToString().ToLowerInvariant()}");
            }
            else
            {
                onLog?.Invoke("Warning", "Parameter RequireSenderAuthenticationEnabled is not supported: change ignored.");
            }
        }

        if (request.AcceptMessagesOnlyFrom != null)
        {
            var acceptParameterName = isDynamicGroup
                ? (SupportsParam("AcceptMessagesOnlyFrom")
                    ? "AcceptMessagesOnlyFrom"
                    : (SupportsParam("AcceptMessagesOnlyFromSendersOrMembers")
                        ? "AcceptMessagesOnlyFromSendersOrMembers"
                        : null))
                : (capabilities.Features.CanSetDistributionGroupAcceptMessagesOnlyFrom ? "AcceptMessagesOnlyFrom" : null);

            if (!string.IsNullOrWhiteSpace(acceptParameterName))
            {
                setParams.Add($"-{acceptParameterName} {ExoRequestSanitizer.FormatStringArrayParameter(request.AcceptMessagesOnlyFrom)}");
            }
            else
            {
                onLog?.Invoke("Warning", "Parameter AcceptMessagesOnlyFrom is not supported: change ignored.");
            }
        }

        if (request.RejectMessagesFrom != null)
        {
            var rejectParameterName = isDynamicGroup
                ? (SupportsParam("RejectMessagesFrom")
                    ? "RejectMessagesFrom"
                    : (SupportsParam("RejectMessagesFromSendersOrMembers")
                        ? "RejectMessagesFromSendersOrMembers"
                        : null))
                : (capabilities.Features.CanSetDistributionGroupRejectMessagesFrom ? "RejectMessagesFrom" : null);

            if (!string.IsNullOrWhiteSpace(rejectParameterName))
            {
                setParams.Add($"-{rejectParameterName} {ExoRequestSanitizer.FormatStringArrayParameter(request.RejectMessagesFrom)}");
            }
            else
            {
                onLog?.Invoke("Warning", "Parameter RejectMessagesFrom is not supported: change ignored.");
            }
        }

        if (setParams.Count == 0)
        {
            return;
        }

        var script = $"{cmdletName} -Identity '{escapedIdentity}' {string.Join(" ", setParams)}";
        onLog?.Invoke("Information", $"Updating group settings for {request.Identity}...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to update group settings: {result.ErrorMessage}");
        }

        onLog?.Invoke("Information", "Group settings updated successfully");
    }

    private static string? NormalizeGroupTypeFilter(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        return NormalizeRequestedGroupType(value) switch
        {
            "All" => null,
            var normalized => normalized
        };
    }

    private static string NormalizeRequestedGroupType(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return GroupTypeDistribution;
        }

        var normalized = value.Trim();
        if (normalized.Equals(GroupTypeDistribution, StringComparison.OrdinalIgnoreCase) ||
            normalized.Equals("DistributionGroup", StringComparison.OrdinalIgnoreCase))
        {
            return GroupTypeDistribution;
        }

        if (normalized.Equals(GroupTypeMailSecurity, StringComparison.OrdinalIgnoreCase) ||
            normalized.Equals("MailEnabledSecurity", StringComparison.OrdinalIgnoreCase) ||
            normalized.Equals("MailUniversalSecurityGroup", StringComparison.OrdinalIgnoreCase))
        {
            return GroupTypeMailSecurity;
        }

        if (normalized.Equals(GroupTypeDynamic, StringComparison.OrdinalIgnoreCase) ||
            normalized.Equals("DynamicDistributionGroup", StringComparison.OrdinalIgnoreCase))
        {
            return GroupTypeDynamic;
        }

        if (normalized.Equals(GroupTypeMicrosoft365, StringComparison.OrdinalIgnoreCase) ||
            normalized.Equals("UnifiedGroup", StringComparison.OrdinalIgnoreCase) ||
            normalized.Equals("Microsoft365Group", StringComparison.OrdinalIgnoreCase))
        {
            return GroupTypeMicrosoft365;
        }

        if (normalized.Equals("All", StringComparison.OrdinalIgnoreCase))
        {
            return "All";
        }

        return GroupTypeDistribution;
    }

    private static string EscapePowerShellString(string? value) => value?.Replace("'", "''") ?? string.Empty;

    private static List<string> ConvertToStringList(object? obj)
    {
        if (obj == null)
        {
            return new List<string>();
        }

        if (obj is object[] array)
        {
            return array.Select(x => x?.ToString() ?? string.Empty).Where(x => !string.IsNullOrEmpty(x)).ToList();
        }

        if (obj is IEnumerable enumerable)
        {
            var list = new List<string>();
            foreach (var item in enumerable)
            {
                var str = item?.ToString();
                if (!string.IsNullOrEmpty(str))
                {
                    list.Add(str);
                }
            }

            return list;
        }

        var single = obj.ToString();
        return string.IsNullOrEmpty(single) ? new List<string>() : new List<string> { single };
    }

    private static List<string> MergeSenderLists(params object?[] values)
    {
        return values
            .SelectMany(ConvertToStringList)
            .Select(v => v?.Trim())
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(v => v, StringComparer.OrdinalIgnoreCase)
            .Cast<string>()
            .ToList();
    }
}
