using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Presentation.ViewModels;

internal sealed class MobileDevicesCapabilityState
{
    private MobileDevicesCapabilityState(
        bool isModuleAvailable,
        bool canLoadPolicies,
        bool canManageAccessState,
        bool canAssignPolicy,
        bool canRemoteWipe,
        string? message)
    {
        IsModuleAvailable = isModuleAvailable;
        CanLoadPolicies = canLoadPolicies;
        CanManageAccessState = canManageAccessState;
        CanAssignPolicy = canAssignPolicy;
        CanRemoteWipe = canRemoteWipe;
        Message = message;
    }

    public bool IsModuleAvailable { get; }

    public bool CanLoadPolicies { get; }

    public bool CanManageAccessState { get; }

    public bool CanAssignPolicy { get; }

    public bool CanRemoteWipe { get; }

    public string? Message { get; }

    public static MobileDevicesCapabilityState Unknown() =>
        new(
            isModuleAvailable: true,
            canLoadPolicies: true,
            canManageAccessState: true,
            canAssignPolicy: true,
            canRemoteWipe: true,
            message: null);

    public static MobileDevicesCapabilityState From(CapabilityMapDto? capabilities)
    {
        if (capabilities?.Features == null)
        {
            return Unknown();
        }

        var features = capabilities.Features;
        if (!features.CanGetMobileDevice)
        {
            var reason = DescribeCmdletAvailability(capabilities, "Get-MobileDevice");
            return new MobileDevicesCapabilityState(
                isModuleAvailable: false,
                canLoadPolicies: false,
                canManageAccessState: false,
                canAssignPolicy: false,
                canRemoteWipe: false,
                message: $"The Mobile Devices module is not available: the current Exchange session does not expose Get-MobileDevice{reason}.");
        }

        var warnings = new List<string>();
        var canLoadPolicies = features.CanGetMobileDeviceMailboxPolicy;
        var canManageAccessState = features.CanGetCasMailbox && features.CanSetCasMailbox;
        var canAssignPolicy = canLoadPolicies && features.CanSetCasMailbox;
        var canRemoteWipe = features.CanClearMobileDevice;

        if (!canLoadPolicies)
        {
            warnings.Add($"mailbox policy not available ({DescribeCmdletAvailability(capabilities, "Get-MobileDeviceMailboxPolicy").TrimStart(':', ' ')})");
        }

        if (!canManageAccessState)
        {
            warnings.Add($"Allow/Block/Quarantine actions are disabled ({DescribeCombinedAvailability(capabilities, "Get-CASMailbox", "Set-CASMailbox")})");
        }

        if (!canRemoteWipe)
        {
            warnings.Add($"remote wipe is disabled ({DescribeCmdletAvailability(capabilities, "Clear-MobileDevice").TrimStart(':', ' ')})");
        }

        return new MobileDevicesCapabilityState(
            isModuleAvailable: true,
            canLoadPolicies: canLoadPolicies,
            canManageAccessState: canManageAccessState,
            canAssignPolicy: canAssignPolicy,
            canRemoteWipe: canRemoteWipe,
            message: warnings.Count == 0
                ? null
                : $"The Mobile Devices module is available with reduced functionality: {string.Join("; ", warnings)}.");
    }

    private static string DescribeCombinedAvailability(CapabilityMapDto capabilities, params string[] cmdletNames)
    {
        var reasons = cmdletNames
            .Select(name => $"{name}{DescribeCmdletAvailability(capabilities, name)}")
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        return string.Join("; ", reasons);
    }

    private static string DescribeCmdletAvailability(CapabilityMapDto capabilities, string cmdletName)
    {
        if (!capabilities.Cmdlets.TryGetValue(cmdletName, out var cmdlet))
        {
            return " (capability not detected)";
        }

        if (cmdlet.IsAvailable)
        {
            return " (cmdlet available)";
        }

        return string.IsNullOrWhiteSpace(cmdlet.UnavailableReason)
            ? " (cmdlet not available)"
            : $": {cmdlet.UnavailableReason}";
    }
}

