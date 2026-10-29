namespace OnlyExo365.Shell.ViewModels;

internal sealed class MailboxSettingsSnapshot
{
    public string PrimarySmtpAddress { get; set; } = string.Empty;
    public List<string> ProxyAddresses { get; set; } = new();
    public bool HiddenFromAddressListsEnabled { get; set; }
    public string ForwardingAddress { get; set; } = string.Empty;
    public string ForwardingSmtpAddress { get; set; } = string.Empty;
    public bool DeliverToMailboxAndForward { get; set; }
    public bool ArchiveEnabled { get; set; }
    public bool LitigationHoldEnabled { get; set; }
    public bool AuditEnabled { get; set; }
    public bool SingleItemRecoveryEnabled { get; set; }
    public bool RetentionHoldEnabled { get; set; }
    public string IssueWarningQuota { get; set; } = string.Empty;
    public string ProhibitSendQuota { get; set; } = string.Empty;
    public string ProhibitSendReceiveQuota { get; set; } = string.Empty;
    public string MaxSendSize { get; set; } = string.Empty;
    public string MaxReceiveSize { get; set; } = string.Empty;
    public string RetentionPolicy { get; set; } = string.Empty;
    public bool OwaEnabled { get; set; }
    public bool ActiveSyncEnabled { get; set; }
    public bool MapiEnabled { get; set; }
    public bool PopEnabled { get; set; }
    public bool ImapEnabled { get; set; }
    public bool SmtpClientAuthenticationDisabled { get; set; }

    public override bool Equals(object? obj)
    {
        if (obj is not MailboxSettingsSnapshot other)
        {
            return false;
        }

        return string.Equals(PrimarySmtpAddress, other.PrimarySmtpAddress, StringComparison.Ordinal)
            && ProxyAddresses.SequenceEqual(other.ProxyAddresses, StringComparer.Ordinal)
            && HiddenFromAddressListsEnabled == other.HiddenFromAddressListsEnabled
            && string.Equals(ForwardingAddress, other.ForwardingAddress, StringComparison.Ordinal)
            && string.Equals(ForwardingSmtpAddress, other.ForwardingSmtpAddress, StringComparison.Ordinal)
            && DeliverToMailboxAndForward == other.DeliverToMailboxAndForward
            && ArchiveEnabled == other.ArchiveEnabled
            && LitigationHoldEnabled == other.LitigationHoldEnabled
            && AuditEnabled == other.AuditEnabled
            && SingleItemRecoveryEnabled == other.SingleItemRecoveryEnabled
            && RetentionHoldEnabled == other.RetentionHoldEnabled
            && string.Equals(IssueWarningQuota, other.IssueWarningQuota, StringComparison.Ordinal)
            && string.Equals(ProhibitSendQuota, other.ProhibitSendQuota, StringComparison.Ordinal)
            && string.Equals(ProhibitSendReceiveQuota, other.ProhibitSendReceiveQuota, StringComparison.Ordinal)
            && string.Equals(MaxSendSize, other.MaxSendSize, StringComparison.Ordinal)
            && string.Equals(MaxReceiveSize, other.MaxReceiveSize, StringComparison.Ordinal)
            && string.Equals(RetentionPolicy, other.RetentionPolicy, StringComparison.Ordinal)
            && OwaEnabled == other.OwaEnabled
            && ActiveSyncEnabled == other.ActiveSyncEnabled
            && MapiEnabled == other.MapiEnabled
            && PopEnabled == other.PopEnabled
            && ImapEnabled == other.ImapEnabled
            && SmtpClientAuthenticationDisabled == other.SmtpClientAuthenticationDisabled;
    }

    public override int GetHashCode()
    {
        var hash = new HashCode();
        hash.Add(PrimarySmtpAddress);
        foreach (var proxyAddress in ProxyAddresses)
        {
            hash.Add(proxyAddress);
        }

        hash.Add(HiddenFromAddressListsEnabled);
        hash.Add(ForwardingAddress);
        hash.Add(ForwardingSmtpAddress);
        hash.Add(DeliverToMailboxAndForward);
        hash.Add(ArchiveEnabled);
        hash.Add(LitigationHoldEnabled);
        hash.Add(AuditEnabled);
        hash.Add(SingleItemRecoveryEnabled);
        hash.Add(RetentionHoldEnabled);
        hash.Add(IssueWarningQuota);
        hash.Add(ProhibitSendQuota);
        hash.Add(ProhibitSendReceiveQuota);
        hash.Add(MaxSendSize);
        hash.Add(MaxReceiveSize);
        hash.Add(RetentionPolicy);
        hash.Add(OwaEnabled);
        hash.Add(ActiveSyncEnabled);
        hash.Add(MapiEnabled);
        hash.Add(PopEnabled);
        hash.Add(ImapEnabled);
        hash.Add(SmtpClientAuthenticationDisabled);
        return hash.ToHashCode();
    }
}

internal sealed class MailboxAutoReplySnapshot
{
    public bool AutoReplyEnabled { get; set; }
    public bool AutoReplyScheduled { get; set; }
    public DateTime? AutoReplyStartDate { get; set; }
    public DateTime? AutoReplyEndDate { get; set; }
    public string AutoReplyInternalMessage { get; set; } = string.Empty;
    public string AutoReplyExternalMessage { get; set; } = string.Empty;
    public string AutoReplyExternalAudience { get; set; } = string.Empty;

    public override bool Equals(object? obj)
    {
        if (obj is not MailboxAutoReplySnapshot other)
        {
            return false;
        }

        return AutoReplyEnabled == other.AutoReplyEnabled
            && AutoReplyScheduled == other.AutoReplyScheduled
            && Nullable.Equals(AutoReplyStartDate, other.AutoReplyStartDate)
            && Nullable.Equals(AutoReplyEndDate, other.AutoReplyEndDate)
            && string.Equals(AutoReplyInternalMessage, other.AutoReplyInternalMessage, StringComparison.Ordinal)
            && string.Equals(AutoReplyExternalMessage, other.AutoReplyExternalMessage, StringComparison.Ordinal)
            && string.Equals(AutoReplyExternalAudience, other.AutoReplyExternalAudience, StringComparison.Ordinal);
    }

    public override int GetHashCode()
    {
        var hash = new HashCode();
        hash.Add(AutoReplyEnabled);
        hash.Add(AutoReplyScheduled);
        hash.Add(AutoReplyStartDate);
        hash.Add(AutoReplyEndDate);
        hash.Add(AutoReplyInternalMessage);
        hash.Add(AutoReplyExternalMessage);
        hash.Add(AutoReplyExternalAudience);
        return hash.ToHashCode();
    }
}

