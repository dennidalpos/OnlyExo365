using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;
using ExchangeAdmin.Infrastructure.Ipc;
using ExchangeAdmin.Presentation.Services;
using ExchangeAdmin.Presentation.ViewModels;

namespace ExchangeAdmin.Tests;

public sealed class MailboxSettingsEditorViewModelTests
{
    [Fact]
    public void BuildUpdateRequest_IncludesPrimarySmtpProxyVisibilityAndQuotaChanges()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        var editor = new MailboxSettingsEditorViewModel(
            new MailboxSettingsWorkerService(),
            shell,
            new CacheService());

        editor.Initialize(new MailboxDetailsDto
        {
            Identity = "shared-mailbox",
            DisplayName = "Shared Mailbox",
            PrimarySmtpAddress = "shared@contoso.com",
            EmailAddresses =
            [
                "SMTP:shared@contoso.com",
                "smtp:alias@contoso.com",
                "x500:/o=Contoso/ou=Legacy/cn=Recipients/cn=shared"
            ],
            Features = new MailboxFeaturesDto
            {
                HiddenFromAddressListsEnabled = false,
                IssueWarningQuota = "47 GB",
                ProhibitSendQuota = "48 GB",
                ProhibitSendReceiveQuota = "49 GB",
                OwaEnabled = true,
                ActiveSyncEnabled = true,
                MapiEnabled = true,
                PopEnabled = false,
                ImapEnabled = false,
                SmtpClientAuthenticationDisabled = false
            }
        });

        editor.PrimarySmtpAddress = "shared-primary@contoso.com";
        editor.ProxyAddressesText = "alias2@contoso.com\r\nX500:/o=Contoso/ou=Legacy/cn=Recipients/cn=shared";
        editor.HiddenFromAddressListsEnabled = true;
        editor.IssueWarningQuota = "50 GB";
        editor.ProhibitSendQuota = "51 GB";
        editor.ProhibitSendReceiveQuota = "52 GB";
        editor.OwaEnabled = false;
        editor.ActiveSyncEnabled = false;
        editor.MapiEnabled = false;
        editor.PopEnabled = true;
        editor.ImapEnabled = true;
        editor.SmtpClientAuthenticationDisabled = true;

        var request = editor.BuildUpdateRequest(
            "shared-mailbox",
            out var settingsChanged,
            out var retentionPolicyChanged,
            out var retentionPolicyOverride);

        Assert.True(settingsChanged);
        Assert.False(retentionPolicyChanged);
        Assert.Equal(string.Empty, retentionPolicyOverride);
        Assert.Equal("shared-primary@contoso.com", request.PrimarySmtpAddress);
        Assert.Equal(
            ["smtp:alias2@contoso.com", "X500:/o=Contoso/ou=Legacy/cn=Recipients/cn=shared"],
            request.ProxyAddresses);
        Assert.True(request.HiddenFromAddressListsEnabled);
        Assert.Equal("50 GB", request.IssueWarningQuota);
        Assert.Equal("51 GB", request.ProhibitSendQuota);
        Assert.Equal("52 GB", request.ProhibitSendReceiveQuota);
        Assert.False(request.OwaEnabled);
        Assert.False(request.ActiveSyncEnabled);
        Assert.False(request.MapiEnabled);
        Assert.True(request.PopEnabled);
        Assert.True(request.ImapEnabled);
        Assert.True(request.SmtpClientAuthenticationDisabled);
    }

    [Fact]
    public void ApplyCapabilities_DisablesUnsupportedCasTogglesAndExposesMessage()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        var editor = new MailboxSettingsEditorViewModel(
            new MailboxSettingsWorkerService(),
            shell,
            new CacheService());

        editor.ApplyCapabilities(new CapabilityMapDto
        {
            Features = new FeatureCapabilitiesDto
            {
                CanGetCasMailbox = true,
                CanSetCasMailbox = true,
                CanSetCasOwaEnabled = true,
                CanSetCasActiveSyncEnabled = false,
                CanSetCasMapiEnabled = true,
                CanSetCasPopEnabled = false,
                CanSetCasImapEnabled = true,
                CanSetCasSmtpClientAuthenticationDisabled = false
            }
        });

        Assert.True(editor.CanEditOwaEnabled);
        Assert.False(editor.CanEditActiveSyncEnabled);
        Assert.True(editor.CanEditMapiEnabled);
        Assert.False(editor.CanEditPopEnabled);
        Assert.True(editor.CanEditImapEnabled);
        Assert.False(editor.CanEditSmtpClientAuthenticationDisabled);
        Assert.NotNull(editor.CasCapabilityMessage);
        Assert.Contains("limitations", editor.CasCapabilityMessage, StringComparison.OrdinalIgnoreCase);
    }

    private sealed class MailboxSettingsWorkerService : TestMailboxesWorkerServiceBase;
}
