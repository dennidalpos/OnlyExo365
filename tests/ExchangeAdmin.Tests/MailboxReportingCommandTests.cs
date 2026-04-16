using System.Collections;
using System.Management.Automation;
using ExchangeAdmin.Worker.PowerShell;

namespace ExchangeAdmin.Tests;

public class MailboxReportingCommandTests
{
    [Fact]
    public void TryReadProgress_SupportsPsCustomObjectOutput()
    {
        var output = CreatePsCustomObject(
            ("EntryType", "Mailbox"),
            ("Index", 4),
            ("TotalCount", 10));

        var success = ExoMailboxReportingCommands.TryReadProgress(output, out var current, out var total);

        Assert.True(success);
        Assert.Equal(4, current);
        Assert.Equal(10, total);
    }

    [Fact]
    public void TryReadProgress_SupportsHashtableOutput()
    {
        var output = PSObject.AsPSObject(new Hashtable
        {
            ["EntryType"] = "Grant",
            ["Index"] = 3,
            ["TotalCount"] = 7
        });

        var success = ExoMailboxReportingCommands.TryReadProgress(output, out var current, out var total);

        Assert.True(success);
        Assert.Equal(3, current);
        Assert.Equal(7, total);
    }

    [Fact]
    public void TryParseMailboxSpaceItem_SupportsPsCustomObjectOutput()
    {
        var output = CreatePsCustomObject(
            ("EntryType", "Mailbox"),
            ("Identity", "shared"),
            ("DisplayName", "Shared mailbox"),
            ("PrimarySmtpAddress", "shared@contoso.com"),
            ("TotalItemSize", "1.23 GB (1320702443 bytes)"),
            ("TotalItemSizeBytes", 1320702443L),
            ("ProhibitSendQuota", "48 GB (51539607552 bytes)"),
            ("ProhibitSendQuotaBytes", 51539607552L),
            ("ProhibitSendReceiveQuota", "49 GB (52613349376 bytes)"),
            ("ProhibitSendReceiveQuotaBytes", 52613349376L),
            ("IssueWarningQuota", "47 GB (50465865728 bytes)"),
            ("IssueWarningQuotaBytes", 50465865728L));

        var item = ExoMailboxReportingCommands.TryParseMailboxSpaceItem(output);

        Assert.NotNull(item);
        Assert.Equal("shared", item!.Identity);
        Assert.Equal("shared@contoso.com", item.PrimarySmtpAddress);
        Assert.Equal(1320702443L, item.TotalItemSizeBytes);
        Assert.Equal(52613349376L, item.ProhibitSendReceiveQuotaBytes);
    }

    [Fact]
    public void TryParseMailboxAccessGrant_SupportsPsCustomObjectOutput()
    {
        var output = CreatePsCustomObject(
            ("EntryType", "Grant"),
            ("User", "delegate@contoso.com"),
            ("MailboxIdentity", "shared"),
            ("MailboxDisplayName", "Shared mailbox"),
            ("MailboxPrimarySmtpAddress", "shared@contoso.com"),
            ("PermissionType", "FullAccess"),
            ("AccessRights", new[] { "FullAccess", "ReadPermission" }));

        var grant = ExoMailboxReportingCommands.TryParseMailboxAccessGrant(output);

        Assert.NotNull(grant);
        Assert.Equal("delegate@contoso.com", grant!.User);
        Assert.Equal("shared@contoso.com", grant.MailboxPrimarySmtpAddress);
        Assert.Equal(["FullAccess", "ReadPermission"], grant.AccessRights);
    }

    private static PSObject CreatePsCustomObject(params (string Name, object? Value)[] properties)
    {
        var obj = new PSObject();
        foreach (var (name, value) in properties)
        {
            obj.Properties.Add(new PSNoteProperty(name, value));
        }

        return obj;
    }
}
