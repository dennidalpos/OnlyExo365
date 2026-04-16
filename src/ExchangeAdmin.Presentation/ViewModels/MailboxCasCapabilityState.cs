using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Presentation.ViewModels;

internal sealed class MailboxCasCapabilityState
{
    private MailboxCasCapabilityState(
        bool canReadSettings,
        bool canEditOwaEnabled,
        bool canEditActiveSyncEnabled,
        bool canEditMapiEnabled,
        bool canEditPopEnabled,
        bool canEditImapEnabled,
        bool canEditSmtpClientAuthenticationDisabled,
        string? message)
    {
        CanReadSettings = canReadSettings;
        CanEditOwaEnabled = canEditOwaEnabled;
        CanEditActiveSyncEnabled = canEditActiveSyncEnabled;
        CanEditMapiEnabled = canEditMapiEnabled;
        CanEditPopEnabled = canEditPopEnabled;
        CanEditImapEnabled = canEditImapEnabled;
        CanEditSmtpClientAuthenticationDisabled = canEditSmtpClientAuthenticationDisabled;
        Message = message;
    }

    public bool CanReadSettings { get; }

    public bool CanEditOwaEnabled { get; }

    public bool CanEditActiveSyncEnabled { get; }

    public bool CanEditMapiEnabled { get; }

    public bool CanEditPopEnabled { get; }

    public bool CanEditImapEnabled { get; }

    public bool CanEditSmtpClientAuthenticationDisabled { get; }

    public string? Message { get; }

    public static MailboxCasCapabilityState Unknown() =>
        new(
            canReadSettings: true,
            canEditOwaEnabled: true,
            canEditActiveSyncEnabled: true,
            canEditMapiEnabled: true,
            canEditPopEnabled: true,
            canEditImapEnabled: true,
            canEditSmtpClientAuthenticationDisabled: true,
            message: null);

    public static MailboxCasCapabilityState From(CapabilityMapDto? capabilities)
    {
        if (capabilities?.Features == null)
        {
            return Unknown();
        }

        var features = capabilities.Features;
        if (!features.CanGetCasMailbox)
        {
            return new MailboxCasCapabilityState(
                canReadSettings: false,
                canEditOwaEnabled: false,
                canEditActiveSyncEnabled: false,
                canEditMapiEnabled: false,
                canEditPopEnabled: false,
                canEditImapEnabled: false,
                canEditSmtpClientAuthenticationDisabled: false,
                message: $"Mailbox protocol configuration is not available: Get-CASMailbox is not supported{DescribeCmdletAvailability(capabilities, "Get-CASMailbox")}.");
        }

        var warnings = new List<string>();
        var canEditOwaEnabled = features.CanSetCasMailbox && features.CanSetCasOwaEnabled;
        var canEditActiveSyncEnabled = features.CanSetCasMailbox && features.CanSetCasActiveSyncEnabled;
        var canEditMapiEnabled = features.CanSetCasMailbox && features.CanSetCasMapiEnabled;
        var canEditPopEnabled = features.CanSetCasMailbox && features.CanSetCasPopEnabled;
        var canEditImapEnabled = features.CanSetCasMailbox && features.CanSetCasImapEnabled;
        var canEditSmtpClientAuthenticationDisabled =
            features.CanSetCasMailbox && features.CanSetCasSmtpClientAuthenticationDisabled;

        AddWarningIfNeeded(warnings, canEditOwaEnabled, "OWA", "OWAEnabled");
        AddWarningIfNeeded(warnings, canEditActiveSyncEnabled, "Exchange ActiveSync", "ActiveSyncEnabled");
        AddWarningIfNeeded(warnings, canEditMapiEnabled, "MAPI/Outlook desktop", "MAPIEnabled");
        AddWarningIfNeeded(warnings, canEditPopEnabled, "POP", "PopEnabled");
        AddWarningIfNeeded(warnings, canEditImapEnabled, "IMAP", "ImapEnabled");
        AddWarningIfNeeded(
            warnings,
            canEditSmtpClientAuthenticationDisabled,
            "SMTP AUTH client",
            "SmtpClientAuthenticationDisabled");

        return new MailboxCasCapabilityState(
            canReadSettings: true,
            canEditOwaEnabled: canEditOwaEnabled,
            canEditActiveSyncEnabled: canEditActiveSyncEnabled,
            canEditMapiEnabled: canEditMapiEnabled,
            canEditPopEnabled: canEditPopEnabled,
            canEditImapEnabled: canEditImapEnabled,
            canEditSmtpClientAuthenticationDisabled: canEditSmtpClientAuthenticationDisabled,
            message: warnings.Count == 0
                ? null
                : $"Mailbox protocol configuration is available with limitations: {string.Join("; ", warnings)}.");

        void AddWarningIfNeeded(List<string> items, bool isAvailable, string label, string parameterName)
        {
            if (!isAvailable)
            {
                items.Add($"{label} is read-only ({parameterName} is not supported)");
            }
        }
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
