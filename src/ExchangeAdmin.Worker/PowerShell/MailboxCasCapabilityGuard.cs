using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

internal static class MailboxCasCapabilityGuard
{
    public static void EnsureMailboxSettingsUpdateAvailable(CapabilityMapDto capabilities, UpdateMailboxSettingsRequest request)
    {
        if (!HasCasMailboxChanges(request))
        {
            return;
        }

        if (!capabilities.Features.CanSetCasMailbox)
        {
            throw new InvalidOperationException(
                $"Mailbox protocol update is not available: Set-CASMailbox is not supported{DescribeCmdletAvailability(capabilities, "Set-CASMailbox")}.");
        }

        EnsureParameterAvailable(capabilities, request.OwaEnabled.HasValue, capabilities.Features.CanSetCasOwaEnabled, "OWAEnabled");
        EnsureParameterAvailable(capabilities, request.ActiveSyncEnabled.HasValue, capabilities.Features.CanSetCasActiveSyncEnabled, "ActiveSyncEnabled");
        EnsureParameterAvailable(capabilities, request.MapiEnabled.HasValue, capabilities.Features.CanSetCasMapiEnabled, "MAPIEnabled");
        EnsureParameterAvailable(capabilities, request.PopEnabled.HasValue, capabilities.Features.CanSetCasPopEnabled, "PopEnabled");
        EnsureParameterAvailable(capabilities, request.ImapEnabled.HasValue, capabilities.Features.CanSetCasImapEnabled, "ImapEnabled");
        EnsureParameterAvailable(
            capabilities,
            request.SmtpClientAuthenticationDisabled.HasValue,
            capabilities.Features.CanSetCasSmtpClientAuthenticationDisabled,
            "SmtpClientAuthenticationDisabled");
    }

    private static bool HasCasMailboxChanges(UpdateMailboxSettingsRequest request)
    {
        return request.OwaEnabled.HasValue
            || request.ActiveSyncEnabled.HasValue
            || request.MapiEnabled.HasValue
            || request.PopEnabled.HasValue
            || request.ImapEnabled.HasValue
            || request.SmtpClientAuthenticationDisabled.HasValue;
    }

    private static void EnsureParameterAvailable(
        CapabilityMapDto capabilities,
        bool isRequested,
        bool isSupported,
        string parameterName)
    {
        if (!isRequested || isSupported)
        {
            return;
        }

        throw new InvalidOperationException(
            $"CAS mailbox update is not available: parameter {parameterName} is not supported by Set-CASMailbox.");
    }

    private static string DescribeCmdletAvailability(CapabilityMapDto capabilities, string cmdletName)
    {
        if (!capabilities.Cmdlets.TryGetValue(cmdletName, out var cmdlet))
        {
            return " (capability not detected)";
        }

        if (cmdlet.IsAvailable)
        {
            return string.Empty;
        }

        return string.IsNullOrWhiteSpace(cmdlet.UnavailableReason)
            ? " (cmdlet not available)"
            : $": {cmdlet.UnavailableReason}";
    }
}
