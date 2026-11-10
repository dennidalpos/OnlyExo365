using OnlyExo365.Contracts.Dtos;

namespace OnlyExo365.Worker.PowerShell;

internal static class MobileDeviceCapabilityGuard
{
    public static void EnsureListingAvailable(CapabilityMapDto capabilities)
    {
        if (!capabilities.Features.CanGetMobileDevice)
        {
            throw new InvalidOperationException(
                $"Mobile Devices are not available: Get-MobileDevice is not supported{DescribeCmdletAvailability(capabilities, "Get-MobileDevice")}.");
        }
    }

    public static void EnsurePoliciesAvailable(CapabilityMapDto capabilities)
    {
        if (!capabilities.Features.CanGetMobileDeviceMailboxPolicy)
        {
            throw new InvalidOperationException(
                $"ActiveSync mailbox policies are not available: Get-MobileDeviceMailboxPolicy is not supported{DescribeCmdletAvailability(capabilities, "Get-MobileDeviceMailboxPolicy")}.");
        }
    }

    public static void EnsureAccessStateManagementAvailable(CapabilityMapDto capabilities)
    {
        if (!capabilities.Features.CanGetCasMailbox || !capabilities.Features.CanSetCasMailbox)
        {
            throw new InvalidOperationException(
                $"Device state management is not available: Get-CASMailbox and Set-CASMailbox are required ({DescribeCombinedAvailability(capabilities, "Get-CASMailbox", "Set-CASMailbox")}).");
        }
    }

    public static void EnsureRemoteWipeAvailable(CapabilityMapDto capabilities)
    {
        if (!capabilities.Features.CanClearMobileDevice)
        {
            throw new InvalidOperationException(
                $"Remote wipe is not available: Clear-MobileDevice is not supported{DescribeCmdletAvailability(capabilities, "Clear-MobileDevice")}.");
        }
    }

    public static void EnsureMailboxPolicyAssignmentAvailable(CapabilityMapDto capabilities)
    {
        if (!capabilities.Features.CanSetCasMailbox)
        {
            throw new InvalidOperationException(
                $"Mailbox policy assignment is not available: Set-CASMailbox is not supported{DescribeCmdletAvailability(capabilities, "Set-CASMailbox")}.");
        }
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

