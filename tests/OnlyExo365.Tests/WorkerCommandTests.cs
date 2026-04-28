using OnlyExo365.Contracts;
using OnlyExo365.Contracts.Diagnostics;
using OnlyExo365.Worker;
using OnlyExo365.Worker.PowerShell;
using System.Collections;
using System.Text.Json;

namespace OnlyExo365.Tests;

public partial class WorkerCommandTests
{
    [Theory]
    [InlineData("SharedMailbox", "SharedMailbox")]
    [InlineData(" sharedmailbox ", "SharedMailbox")]
    [InlineData("InvalidMailbox", null)]
    [InlineData("", null)]
    public void NormalizeMailboxRecipientTypeDetails_UsesAllowList(string input, string? expected)
    {
        var normalized = ExoRequestSanitizer.NormalizeMailboxRecipientTypeDetails(input);

        Assert.Equal(expected, normalized);
    }

    [Theory]
    [InlineData("PrimarySmtpAddress", "PrimarySmtpAddress")]
    [InlineData(" recipienttypedetails ", "RecipientTypeDetails")]
    [InlineData("BadProperty", "DisplayName")]
    [InlineData("", "DisplayName")]
    public void NormalizeSortProperties_FallBackToDisplayName(string input, string expected)
    {
        Assert.Equal(expected, ExoRequestSanitizer.NormalizeMailboxSortProperty(input));
        Assert.Equal(expected == "PrimarySmtpAddress" || expected == "RecipientTypeDetails" ? expected : "DisplayName",
            ExoRequestSanitizer.NormalizeGroupSortProperty(input));
    }

    [Fact]
    public void FormatStringArrayParameter_EscapesAndDropsEmptyValues()
    {
        var formatted = ExoRequestSanitizer.FormatStringArrayParameter(new[] { " user1@contoso.com ", "", "o'hara@contoso.com" });

        Assert.Equal("@('user1@contoso.com', 'o''hara@contoso.com')", formatted);
    }

    [Fact]
    public void BuildConnectExchangeCommand_UsesSupportedValuesAndEscapesInputs()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            ExchangeEnvironmentName = "UnsupportedCloud",
            ExchangeOrganization = "tenant.onmicrosoft.com",
            DelegatedOrganization = "delegate'onmicrosoft.com",
            UserPrincipalNameHint = "admin'o@contoso.com"
        };

        var command = ExchangeCommandBuilder.BuildConnectExchangeCommand(configuration);

        Assert.StartsWith("Connect-ExchangeOnline -ShowBanner:$false", command, StringComparison.Ordinal);
        Assert.DoesNotContain("-ExchangeEnvironmentName", command, StringComparison.Ordinal);
        Assert.Contains("-Organization 'tenant.onmicrosoft.com'", command, StringComparison.Ordinal);
        Assert.Contains("-DelegatedOrganization 'delegate''onmicrosoft.com'", command, StringComparison.Ordinal);
        Assert.Contains("-UserPrincipalName 'admin''o@contoso.com'", command, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildConnectExchangeCommand_BuildsAppCertificateFlow()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            AuthenticationMode = ExchangeAuthenticationMode.AppCertificate,
            ExchangeEnvironmentName = "O365Default",
            ExchangeOrganization = "tenant.onmicrosoft.com",
            ApplicationId = "app-id",
            CertificateThumbprint = "ABC123"
        };

        var command = ExchangeCommandBuilder.BuildConnectExchangeCommand(configuration);

        Assert.Contains("-AppId 'app-id'", command, StringComparison.Ordinal);
        Assert.Contains("-CertificateThumbprint 'ABC123'", command, StringComparison.Ordinal);
        Assert.Contains("-Organization 'tenant.onmicrosoft.com'", command, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildConnectExchangeCommand_BuildsManagedIdentityFlow()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            AuthenticationMode = ExchangeAuthenticationMode.ManagedIdentity,
            ExchangeEnvironmentName = "O365Default",
            ExchangeOrganization = "tenant.onmicrosoft.com",
            ManagedIdentityAccountId = "mi-client-id"
        };

        var command = ExchangeCommandBuilder.BuildConnectExchangeCommand(configuration);

        Assert.Contains("-ManagedIdentity", command, StringComparison.Ordinal);
        Assert.Contains("-ManagedIdentityAccountId 'mi-client-id'", command, StringComparison.Ordinal);
        Assert.Contains("-Organization 'tenant.onmicrosoft.com'", command, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildConnectExchangeCommand_BuildsDeviceCodeFlow()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            AuthenticationMode = ExchangeAuthenticationMode.DeviceCode,
            ExchangeEnvironmentName = "O365Default",
            ExchangeOrganization = "tenant.onmicrosoft.com",
            DelegatedOrganization = "delegate.onmicrosoft.com",
            UserPrincipalNameHint = "admin@contoso.com"
        };

        var command = ExchangeCommandBuilder.BuildConnectExchangeCommand(configuration);

        Assert.Contains("-Device", command, StringComparison.Ordinal);
        Assert.Contains("-Organization 'tenant.onmicrosoft.com'", command, StringComparison.Ordinal);
        Assert.Contains("-DelegatedOrganization 'delegate.onmicrosoft.com'", command, StringComparison.Ordinal);
        Assert.Contains("-UserPrincipalName 'admin@contoso.com'", command, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildConnectExchangeCommand_ThrowsForInvalidExchangeOrganization()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            AuthenticationMode = ExchangeAuthenticationMode.Interactive,
            ExchangeOrganization = "tenant with spaces"
        };

        var exception = Assert.Throws<InvalidOperationException>(() => ExchangeCommandBuilder.BuildConnectExchangeCommand(configuration));

        Assert.Contains("ExchangeOrganization must be a tenant domain like 'contoso.onmicrosoft.com' or a tenant GUID.", exception.Message);
    }

    [Fact]
    public void BuildConnectGraphCommand_BuildsAppCertificateFlowUsingCertificateSubject()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            AuthenticationMode = ExchangeAuthenticationMode.AppCertificate,
            ApplicationId = "app-id",
            GraphTenantId = "contoso.onmicrosoft.com",
            CertificateSubjectName = "CN=OnlyExo365 Worker"
        };

        var command = GraphCommandBuilder.BuildConnectGraphCommand(configuration);

        Assert.Contains("-ClientId 'app-id'", command, StringComparison.Ordinal);
        Assert.Contains("-TenantId 'contoso.onmicrosoft.com'", command, StringComparison.Ordinal);
        Assert.Contains("-CertificateSubjectName 'CN=OnlyExo365 Worker'", command, StringComparison.Ordinal);
        Assert.DoesNotContain("-Scopes", command, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildConnectGraphCommand_ThrowsForInvalidGraphTenantId()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            AuthenticationMode = ExchangeAuthenticationMode.Interactive,
            GraphTenantId = "tenant id with spaces",
            GraphScopes = new List<string> { "User.Read.All" }
        };

        var exception = Assert.Throws<InvalidOperationException>(() => GraphCommandBuilder.BuildConnectGraphCommand(configuration));

        Assert.Contains("GraphTenantId must be a tenant domain like 'contoso.onmicrosoft.com' or a tenant GUID.", exception.Message);
    }

    [Fact]
    public void BuildConnectGraphCommand_BuildsManagedIdentityFlow()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            AuthenticationMode = ExchangeAuthenticationMode.ManagedIdentity,
            ManagedIdentityAccountId = "mi-client-id"
        };

        var command = GraphCommandBuilder.BuildConnectGraphCommand(configuration);

        Assert.Equal("Connect-MgGraph -Identity -ClientId 'mi-client-id' -ContextScope Process -NoWelcome", command);
    }

    [Fact]
    public void BuildConnectComplianceCommand_BuildsDelegatedFlow()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            AuthenticationMode = ExchangeAuthenticationMode.Interactive,
            ExchangeOrganization = "tenant.onmicrosoft.com",
            UserPrincipalNameHint = "admin@contoso.com"
        };

        var command = ExoCommands.BuildConnectComplianceCommand(configuration);

        Assert.Contains("Connect-IPPSSession", command, StringComparison.Ordinal);
        Assert.Contains("-Organization 'tenant.onmicrosoft.com'", command, StringComparison.Ordinal);
        Assert.Contains("-UserPrincipalName 'admin@contoso.com'", command, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildConnectComplianceCommand_BuildsDeviceCodeFlow()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            AuthenticationMode = ExchangeAuthenticationMode.DeviceCode,
            ExchangeOrganization = "tenant.onmicrosoft.com",
            DelegatedOrganization = "delegate.onmicrosoft.com",
            UserPrincipalNameHint = "admin@contoso.com"
        };

        var command = ExoCommands.BuildConnectComplianceCommand(configuration);

        Assert.Contains("Connect-IPPSSession", command, StringComparison.Ordinal);
        Assert.Contains("-Device", command, StringComparison.Ordinal);
        Assert.Contains("-DelegatedOrganization 'delegate.onmicrosoft.com'", command, StringComparison.Ordinal);
        Assert.Contains("-UserPrincipalName 'admin@contoso.com'", command, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildConnectComplianceCommand_BuildsAppCertificateFlow()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            AuthenticationMode = ExchangeAuthenticationMode.AppCertificate,
            ExchangeOrganization = "tenant.onmicrosoft.com",
            ApplicationId = "app-id",
            CertificateThumbprint = "ABC123"
        };

        var command = ExoCommands.BuildConnectComplianceCommand(configuration);

        Assert.Contains("-AppId 'app-id'", command, StringComparison.Ordinal);
        Assert.Contains("-Organization 'tenant.onmicrosoft.com'", command, StringComparison.Ordinal);
        Assert.Contains("-CertificateThumbprint 'ABC123'", command, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildConnectComplianceCommand_BuildsManagedIdentityFlow()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            AuthenticationMode = ExchangeAuthenticationMode.ManagedIdentity,
            ExchangeOrganization = "tenant.onmicrosoft.com",
            ManagedIdentityAccountId = "mi-client-id"
        };

        var command = ExoCommands.BuildConnectComplianceCommand(configuration);

        Assert.Contains("-ManagedIdentity", command, StringComparison.Ordinal);
        Assert.Contains("-Organization 'tenant.onmicrosoft.com'", command, StringComparison.Ordinal);
        Assert.Contains("-ManagedIdentityAccountId 'mi-client-id'", command, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildConnectComplianceCommand_ThrowsForInvalidExchangeOrganization()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            AuthenticationMode = ExchangeAuthenticationMode.Interactive,
            ExchangeOrganization = "tenant with spaces"
        };

        var exception = Assert.Throws<InvalidOperationException>(() => ExoCommands.BuildConnectComplianceCommand(configuration));

        Assert.Contains("ExchangeOrganization must be a tenant domain like 'contoso.onmicrosoft.com' or a tenant GUID.", exception.Message);
    }

    [Fact]
    public void BuildCreateComplianceSearchScript_UsesLocationsAndOptionalQuery()
    {
        var request = new CreateComplianceSearchRequest
        {
            Name = "Mailbox purge prep",
            CaseName = "Incident 2026-03",
            ExchangeLocations = new List<string> { "user1@contoso.com", "shared@contoso.com" },
            ContentMatchQuery = "kind:email AND subject:\"invoice\""
        };

        var script = ExoCommands.BuildCreateComplianceSearchScript(request);

        Assert.Contains("New-ComplianceSearch @params -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("Name = 'Mailbox purge prep'", script, StringComparison.Ordinal);
        Assert.Contains("ExchangeLocation = @('user1@contoso.com', 'shared@contoso.com')", script, StringComparison.Ordinal);
        Assert.Contains("$params['Case'] = 'Incident 2026-03'", script, StringComparison.Ordinal);
        Assert.Contains("$params['ContentMatchQuery'] = 'kind:email AND subject:\"invoice\"'", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildInvokeComplianceActionScript_SupportsPurgeAndHold()
    {
        var purgeScript = ExoCommands.BuildInvokeComplianceActionScript(
            new InvokeComplianceActionRequest
            {
                SearchName = "Search-01",
                ActionType = "Purge",
                PurgeType = "SoftDelete"
            },
            new List<string> { "user1@contoso.com" },
            "from:\"alerts@contoso.com\"");

        Assert.Contains("New-ComplianceSearchAction -SearchName 'Search-01' -Purge -PurgeType 'SoftDelete'", purgeScript, StringComparison.Ordinal);

        var holdScript = ExoCommands.BuildInvokeComplianceActionScript(
            new InvokeComplianceActionRequest
            {
                SearchName = "Search-01",
                ActionType = "Hold",
                CaseName = "Case-01",
                HoldName = "Search-01 Hold"
            },
            new List<string> { "user1@contoso.com", "user2@contoso.com" },
            "kind:email");

        Assert.Contains("New-CaseHoldPolicy -Name $holdName -Case $caseName -ExchangeLocation $locations -ErrorAction Stop", holdScript, StringComparison.Ordinal);
        Assert.Contains("New-CaseHoldRule -Name 'Search-01 Hold Rule' -Policy $holdName -ContentMatchQuery 'kind:email' -ErrorAction Stop", holdScript, StringComparison.Ordinal);
        Assert.Contains("$locations = @('user1@contoso.com', 'user2@contoso.com')", holdScript, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildUpsertContactCommand_DoesNotInlinePlainTextPassword()
    {
        var request = new UpsertContactRequest
        {
            ContactKind = "MailUser",
            DisplayName = "Mario Rossi",
            Name = "Mario Rossi",
            Alias = "mrossi",
            PrimarySmtpAddress = "mrossi@contoso.com",
            ExternalEmailAddress = "mario.rossi@example.com",
            UserPrincipalName = "mrossi@contoso.com",
            Password = "Sup3rSecret!"
        };

        var command = ExoCommands.BuildUpsertContactCommand(request);

        Assert.DoesNotContain(request.Password, command.Script, StringComparison.Ordinal);
        Assert.Contains("ConvertTo-SecureString $PlainTextPassword -AsPlainText -Force", command.Script, StringComparison.Ordinal);
        Assert.NotNull(command.Parameters);
        Assert.Equal(request.Password, command.Parameters!["PlainTextPassword"]);
    }

    [Fact]
    public void BuildCreateMailboxCommand_DoesNotInlinePlainTextPassword()
    {
        var request = new CreateMailboxRequest
        {
            DisplayName = "Mario Rossi",
            Alias = "mrossi",
            PrimarySmtpAddress = "mrossi@contoso.com",
            MailboxType = "User",
            Password = "Sup3rSecret!"
        };

        var command = ExoMailboxCommands.BuildCreateMailboxCommand(request);

        Assert.DoesNotContain(request.Password, command.Script, StringComparison.Ordinal);
        Assert.Contains("ConvertTo-SecureString $PlainTextPassword -AsPlainText -Force", command.Script, StringComparison.Ordinal);
        Assert.NotNull(command.Parameters);
        Assert.Equal(request.Password, command.Parameters!["PlainTextPassword"]);
    }

    [Fact]
    public void BuildUpsertMigrationEndpointCommand_DoesNotInlinePassword()
    {
        var request = new UpsertMigrationEndpointRequest
        {
            Name = "Remote Move",
            EndpointType = "ExchangeRemoteMove",
            RemoteServer = "mail.contoso.com",
            Username = "admin@contoso.com",
            Password = "Sup3rSecret!"
        };

        var command = ExoCommands.BuildUpsertMigrationEndpointCommand(request);

        Assert.DoesNotContain(request.Password, command.Script, StringComparison.Ordinal);
        Assert.Contains("ConvertTo-SecureString $MigrationPassword -AsPlainText -Force", command.Script, StringComparison.Ordinal);
        Assert.NotNull(command.Parameters);
        Assert.Equal(request.Password, command.Parameters!["MigrationPassword"]);
    }

    [Fact]
    public void BuildTestMigrationEndpointCommand_UsesEndpointWhenRequested()
    {
        var request = new TestMigrationEndpointRequest
        {
            Identity = "endpoint-01",
            UseExistingEndpoint = true,
            EndpointType = "ExchangeRemoteMove"
        };

        var command = ExoCommands.BuildTestMigrationEndpointCommand(request);

        Assert.Contains("Test-MigrationServerAvailability -Endpoint 'endpoint-01'", command.Script, StringComparison.Ordinal);
        Assert.Null(command.Parameters);
    }

    [Fact]
    public void BuildCreateMigrationBatchCommand_UsesCsvDataAndEndpoint()
    {
        var request = new CreateMigrationBatchRequest
        {
            Name = "Batch 01",
            BatchType = "IMAP",
            EndpointIdentity = "endpoint-01",
            CsvFilePath = "C:\\temp\\migration.csv",
            NotificationEmails = ["ops@contoso.com"],
            AutoStart = true
        };

        var command = ExoCommands.BuildCreateMigrationBatchCommand(request);

        Assert.Contains("CSVData = $csvBytes", command.Script, StringComparison.Ordinal);
        Assert.Contains("$params['SourceEndpoint'] = 'endpoint-01'", command.Script, StringComparison.Ordinal);
        Assert.Contains("$params['NotificationEmails'] = @('ops@contoso.com')", command.Script, StringComparison.Ordinal);
        Assert.NotNull(command.Parameters);
        Assert.Equal(request.CsvFilePath, command.Parameters!["CsvFilePath"]);
    }

    [Fact]
    public void BuildUpsertPublicFolderScript_UsesPathChangeWhenParentChanges()
    {
        var script = ExoCommands.BuildUpsertPublicFolderScript(new UpsertPublicFolderRequest
        {
            Identity = "\\Operations\\Queue",
            Name = "Queue",
            ParentPath = "\\Archive",
            MailEnabled = true,
            Alias = "queue-archive",
            PrimarySmtpAddress = "queue-archive@contoso.com",
            HiddenFromAddressListsEnabled = true
        });

        Assert.Contains("Set-PublicFolder -Identity $targetIdentity -Path $parentPath -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("Enable-MailPublicFolder -Identity $targetIdentity -Alias $alias -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("Set-MailPublicFolder @setParams | Out-Null", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildSetPublicFolderClientPermissionScript_ReplacesPermissionOnModify()
    {
        var script = ExoCommands.BuildSetPublicFolderClientPermissionScript(new SetPublicFolderClientPermissionRequest
        {
            Identity = "\\Operations\\Queue",
            User = "delegate@contoso.com",
            Action = PermissionAction.Modify,
            AccessRights = ["PublishingEditor"]
        });

        Assert.Contains("Remove-PublicFolderClientPermission -Identity $identity -User $user -Confirm:$false -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("Add-PublicFolderClientPermission -Identity $identity -User $user -AccessRights $accessRights -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("$accessRights = @('PublishingEditor')", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildRemovePublicFolderScript_DisablesMailBeforeDeleteAndUsesRecurse()
    {
        var script = ExoCommands.BuildRemovePublicFolderScript(new RemovePublicFolderRequest
        {
            Identity = "\\Operations\\Queue",
            Recursive = true
        });

        Assert.Contains("Disable-MailPublicFolder -Identity $identity -Confirm:$false -ErrorAction Stop | Out-Null", script, StringComparison.Ordinal);
        Assert.Contains("$removeParams['Recurse'] = $true", script, StringComparison.Ordinal);
        Assert.Contains("Remove-PublicFolder @removeParams | Out-Null", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildGetMailboxFolderPermissionsScript_UsesFolderScopeResolverForCalendar()
    {
        var script = ExoCommands.BuildGetMailboxFolderPermissionsScript(new GetMailboxFolderPermissionsRequest
        {
            MailboxIdentity = "shared@contoso.com",
            FolderPath = "Calendar\\Team"
        });

        Assert.Contains("Get-MailboxFolderStatistics -Identity $MailboxIdentity -FolderScope $scope -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("Get-MailboxFolderPermission -Identity $resolvedFolderIdentity -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("$folderPath = 'Calendar\\Team'", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildGetMailboxFolderPermissionsScript_MapsItalianCalendarAliasToCalendarScope()
    {
        var script = ExoCommands.BuildGetMailboxFolderPermissionsScript(new GetMailboxFolderPermissionsRequest
        {
            MailboxIdentity = "shared@contoso.com",
            FolderPath = "Calendario"
        });

        Assert.Contains("function Normalize-MailboxFolderSegment", script, StringComparison.Ordinal);
        Assert.Contains("'calendario' = 'Calendar'", script, StringComparison.Ordinal);
        Assert.Contains("$normalizedRoot = Normalize-MailboxFolderSegment $rootSegment", script, StringComparison.Ordinal);
        Assert.Contains("Get-MailboxFolderStatistics -Identity $MailboxIdentity -FolderScope $scope -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("$resolvedRootPath = $scopeEntry.FolderPath.ToString().Trim('/')", script, StringComparison.Ordinal);
        Assert.Contains("$resolvedRoot = \"${MailboxIdentity}:\\$resolvedRootPath\"", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildGetMailboxFolderPermissionsScript_UsesSafeMailboxIdentityInterpolationInFallback()
    {
        var script = ExoCommands.BuildGetMailboxFolderPermissionsScript(new GetMailboxFolderPermissionsRequest
        {
            MailboxIdentity = "shared@contoso.com",
            FolderPath = "Custom"
        });

        Assert.Contains("return \"${MailboxIdentity}:\\$normalizedPath\"", script, StringComparison.Ordinal);
        Assert.DoesNotContain("return \"$MailboxIdentity:\\$normalizedPath\"", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildSetMailboxFolderPermissionScript_UsesSetCmdletForModify()
    {
        var script = ExoCommands.BuildSetMailboxFolderPermissionScript(new SetMailboxFolderPermissionRequest
        {
            MailboxIdentity = "shared@contoso.com",
            FolderPath = "Calendar",
            User = "delegate@contoso.com",
            Action = PermissionAction.Modify,
            AccessRights = ["Editor"]
        });

        Assert.Contains("$action = 'Modify'", script, StringComparison.Ordinal);
        Assert.Contains("$accessRights = @('Editor')", script, StringComparison.Ordinal);
        Assert.Contains("Set-MailboxFolderPermission -Identity $resolvedFolderIdentity -User $user -AccessRights $accessRights -Confirm:$false -ErrorAction Stop", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildSetMailboxFolderPermissionScript_MapsItalianCalendarAliasToCalendarScope()
    {
        var script = ExoCommands.BuildSetMailboxFolderPermissionScript(new SetMailboxFolderPermissionRequest
        {
            MailboxIdentity = "shared@contoso.com",
            FolderPath = "Calendario",
            User = "delegate@contoso.com",
            Action = PermissionAction.Add,
            AccessRights = ["Editor"]
        });

        Assert.Contains("$folderPath = 'Calendario'", script, StringComparison.Ordinal);
        Assert.Contains("'calendario' = 'Calendar'", script, StringComparison.Ordinal);
        Assert.Contains("$resolvedFolderIdentity = Resolve-MailboxFolderIdentity -MailboxIdentity $mailboxIdentity -FolderPath $folderPath", script, StringComparison.Ordinal);
        Assert.Contains("Add-MailboxFolderPermission -Identity $resolvedFolderIdentity -User $user -AccessRights $accessRights -Confirm:$false -ErrorAction Stop", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildSetMailboxSettingsScript_ComposesArchiveAndSetMailboxChanges()
    {
        var request = new UpdateMailboxSettingsRequest
        {
            Identity = "shared'o",
            ArchiveEnabled = false,
            PrimarySmtpAddress = "shared@contoso.com",
            ProxyAddresses = ["smtp:alias@contoso.com", "X500:/o=Legacy/ou=Exchange Administrative Group/cn=Recipients/cn=shared"],
            HiddenFromAddressListsEnabled = true,
            LitigationHoldEnabled = true,
            AuditEnabled = false,
            IssueWarningQuota = "48 GB",
            ProhibitSendQuota = "49 GB",
            ProhibitSendReceiveQuota = "50 GB",
            ForwardingSmtpAddress = "forward@contoso.com",
            DeliverToMailboxAndForward = true,
            MaxSendSize = "50 MB",
            OwaEnabled = false,
            ActiveSyncEnabled = true,
            MapiEnabled = true,
            PopEnabled = false,
            ImapEnabled = false,
            SmtpClientAuthenticationDisabled = true
        };

        var script = ExoMailboxCommands.BuildSetMailboxSettingsScript(request);

        Assert.Contains("Disable-Mailbox -Identity 'shared''o' -Archive -Confirm:$false", script, StringComparison.Ordinal);
        Assert.Contains("Set-Mailbox -Identity 'shared''o'", script, StringComparison.Ordinal);
        Assert.Contains("-EmailAddresses @('SMTP:shared@contoso.com', 'smtp:alias@contoso.com', 'X500:/o=Legacy/ou=Exchange Administrative Group/cn=Recipients/cn=shared')", script, StringComparison.Ordinal);
        Assert.Contains("-HiddenFromAddressListsEnabled $true", script, StringComparison.Ordinal);
        Assert.Contains("-LitigationHoldEnabled $true", script, StringComparison.Ordinal);
        Assert.Contains("-AuditEnabled $false", script, StringComparison.Ordinal);
        Assert.Contains("-IssueWarningQuota '48 GB'", script, StringComparison.Ordinal);
        Assert.Contains("-ProhibitSendQuota '49 GB'", script, StringComparison.Ordinal);
        Assert.Contains("-ProhibitSendReceiveQuota '50 GB'", script, StringComparison.Ordinal);
        Assert.Contains("-ForwardingSmtpAddress 'forward@contoso.com'", script, StringComparison.Ordinal);
        Assert.Contains("-DeliverToMailboxAndForward $true", script, StringComparison.Ordinal);
        Assert.Contains("-MaxSendSize '50 MB'", script, StringComparison.Ordinal);
        Assert.Contains("Set-CASMailbox -Identity 'shared''o'", script, StringComparison.Ordinal);
        Assert.Contains("-OWAEnabled $false", script, StringComparison.Ordinal);
        Assert.Contains("-ActiveSyncEnabled $true", script, StringComparison.Ordinal);
        Assert.Contains("-MAPIEnabled $true", script, StringComparison.Ordinal);
        Assert.Contains("-PopEnabled $false", script, StringComparison.Ordinal);
        Assert.Contains("-ImapEnabled $false", script, StringComparison.Ordinal);
        Assert.Contains("-SmtpClientAuthenticationDisabled $true", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildRestoreMailboxCommand_UsesParametersForRuntimeValues()
    {
        var request = new RestoreMailboxRequest
        {
            SourceIdentity = "soft'user",
            TargetMailbox = "target@contoso.com",
            AllowLegacyDnMismatch = true
        };

        var command = ExoMailboxCommands.BuildRestoreMailboxCommand(request);

        Assert.DoesNotContain(request.SourceIdentity, command.Script, StringComparison.Ordinal);
        Assert.DoesNotContain(request.TargetMailbox, command.Script, StringComparison.Ordinal);
        Assert.Contains("param(", command.Script, StringComparison.Ordinal);
        Assert.Contains("Undo-SoftDeletedMailbox -Identity $identity -AllowLegacyDNMismatch", command.Script, StringComparison.Ordinal);
        Assert.Equal(request.SourceIdentity, command.Parameters["SourceIdentity"]);
        Assert.Equal(request.TargetMailbox, command.Parameters["TargetMailbox"]);
        Assert.Equal(request.AllowLegacyDnMismatch, command.Parameters["AllowLegacyDnMismatch"]);
    }

    [Fact]
    public void BuildRestoreMailboxCommand_UsesInactiveScenarioForInactiveMailboxLookup()
    {
        var command = ExoMailboxCommands.BuildRestoreMailboxCommand(new RestoreMailboxRequest
        {
            SourceIdentity = "inactive-user",
            TargetMailbox = "target@contoso.com"
        });

        Assert.Contains("Get-Mailbox -InactiveMailboxOnly -Identity $SourceIdentity -ErrorAction Stop", command.Script, StringComparison.Ordinal);
        Assert.Contains("$result.Scenario = 'Inactive'", command.Script, StringComparison.Ordinal);
        Assert.DoesNotContain("$result.Scenario = 'HardDeleted'", command.Script, StringComparison.Ordinal);
        Assert.Contains("Inactive mailbox restore request submitted", command.Script, StringComparison.Ordinal);
    }

    [Fact]
    public void ToDeletedMailboxItem_ParsesInactiveDeletionType()
    {
        var hash = new Hashtable
        {
            ["Identity"] = "inactive-user",
            ["DisplayName"] = "Inactive User",
            ["PrimarySmtpAddress"] = "inactive@contoso.com",
            ["RecipientTypeDetails"] = "UserMailbox",
            ["DeletionType"] = "Inactive"
        };

        var item = ExoMailboxMapper.ToDeletedMailboxItem(hash);

        Assert.Equal(DeletedMailboxDeletionType.Inactive, item.DeletionType);
    }

    [Fact]
    public void BuildGetMobileDeviceMailboxPoliciesScript_DoesNotUseUnsupportedResultSizeParameter()
    {
        var script = ExoCommands.BuildGetMobileDeviceMailboxPoliciesScript();

        Assert.Contains("Get-MobileDeviceMailboxPolicy -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.DoesNotContain("-ResultSize", script, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0, 50, 51)]
    [InlineData(10, 25, 36)]
    [InlineData(25, 10, 36)]
    public void CalculateMobileDevicePageWindowSize_LimitsQueryToSkipPlusPageSizePlusOne(int skip, int pageSize, int expected)
    {
        var result = ExoCommands.CalculateMobileDevicePageWindowSize(skip, pageSize);

        Assert.Equal(expected, result);
    }

    [Fact]
    public void BuildGetMobileDevicesScript_UsesUnlimitedScanToKeepExactCounts()
    {
        var script = ExoCommands.BuildGetMobileDevicesScript(
            skip: 10,
            pageSize: 25,
            escapedSearch: string.Empty,
            escapedAccessState: string.Empty,
            sortProperty: "UserDisplayName",
            sortDirection: string.Empty);

        Assert.Contains("Get-MobileDevice -ResultSize Unlimited -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.DoesNotContain("$pageWindowSize =", script, StringComparison.Ordinal);
        Assert.Contains("| Sort-Object UserDisplayName", script, StringComparison.Ordinal);
        Assert.Contains("$pagedItems = @($allItems | Select-Object -Skip 10 -First 25)", script, StringComparison.Ordinal);
        Assert.DoesNotContain("Get-MobileDeviceStatistics", script, StringComparison.Ordinal);
        Assert.DoesNotContain("Get-CASMailbox", script, StringComparison.Ordinal);
        Assert.Contains("HasMore = $hasMore", script, StringComparison.Ordinal);
        Assert.Contains("IsTotalCountExact = $true", script, StringComparison.Ordinal);
        Assert.Contains("$totalCount = $allItems.Count", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildGetMobileDevicesScript_SeedsUserPrincipalNameFromAvailableDeviceProperties()
    {
        var script = ExoCommands.BuildGetMobileDevicesScript(
            skip: 0,
            pageSize: 50,
            escapedSearch: string.Empty,
            escapedAccessState: string.Empty,
            sortProperty: "UserDisplayName",
            sortDirection: string.Empty);

        Assert.Contains("function Resolve-UserPrincipalName", script, StringComparison.Ordinal);
        Assert.Contains("$userPrincipalName = Resolve-UserPrincipalName $device $mailboxIdentity", script, StringComparison.Ordinal);
        Assert.Contains("UserPrincipalName = $userPrincipalName", script, StringComparison.Ordinal);
        Assert.DoesNotContain("UserPrincipalName = $null", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildGetMobileDevicesScript_PrefersResolvedMailboxIdentityWhenUpnIsAvailable()
    {
        var script = ExoCommands.BuildGetMobileDevicesScript(
            skip: 0,
            pageSize: 50,
            escapedSearch: string.Empty,
            escapedAccessState: string.Empty,
            sortProperty: "UserDisplayName",
            sortDirection: string.Empty);

        Assert.Contains("function Resolve-PreferredMailboxIdentity", script, StringComparison.Ordinal);
        Assert.Contains("$resolvedMailboxIdentity = Resolve-PreferredMailboxIdentity $mailboxIdentity $userPrincipalName", script, StringComparison.Ordinal);
        Assert.Contains("MailboxIdentity = $resolvedMailboxIdentity", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildGetMobileDeviceDetailsScript_LoadsSingleDeviceOnly()
    {
        var script = ExoMobileDeviceCommands.BuildGetMobileDeviceDetailsScript("device-01");

        Assert.Contains("Get-MobileDevice -Identity 'device-01' -ErrorAction Stop | Select-Object -First 1", script, StringComparison.Ordinal);
        Assert.DoesNotContain("-ResultSize", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildGetMobileDeviceDetailsScript_SeedsUserPrincipalNameFromAvailableDeviceProperties()
    {
        var script = ExoMobileDeviceCommands.BuildGetMobileDeviceDetailsScript("device-01");

        Assert.Contains("function Resolve-UserPrincipalName", script, StringComparison.Ordinal);
        Assert.Contains("$userPrincipalName = Resolve-UserPrincipalName $device $mailboxIdentity", script, StringComparison.Ordinal);
        Assert.Contains("UserPrincipalName = $userPrincipalName", script, StringComparison.Ordinal);
        Assert.DoesNotContain("UserPrincipalName = $null", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildGetMobileDeviceDetailsScript_PrefersResolvedMailboxIdentityWhenUpnIsAvailable()
    {
        var script = ExoMobileDeviceCommands.BuildGetMobileDeviceDetailsScript("device-01");

        Assert.Contains("function Resolve-PreferredMailboxIdentity", script, StringComparison.Ordinal);
        Assert.Contains("$resolvedMailboxIdentity = Resolve-PreferredMailboxIdentity $mailboxIdentity $userPrincipalName", script, StringComparison.Ordinal);
        Assert.Contains("MailboxIdentity = $resolvedMailboxIdentity", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildGetMessageTraceScript_EnumeratesAllV2PagesToCalculateExactTotal()
    {
        var script = ExoMessageTraceCommands.BuildGetMessageTraceScript(new GetMessageTraceRequest
        {
            StartDate = new DateTime(2026, 3, 1, 0, 0, 0, DateTimeKind.Utc),
            EndDate = new DateTime(2026, 3, 14, 23, 59, 59, DateTimeKind.Utc),
            SenderAddress = "sender@contoso.com",
            RecipientAddress = "recipient@contoso.com",
            Page = 2,
            PageSize = 100
        });

        Assert.Contains("$requestedStartIndex = (($page - 1) * $pageSize)", script, StringComparison.Ordinal);
        Assert.Contains("Get-MessageTraceV2 @batchParams -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("ResultSize = 5000", script, StringComparison.Ordinal);
        Assert.Contains("$batchParams['StartingRecipientAddress'] = $cursorRecipientAddress", script, StringComparison.Ordinal);
        Assert.Contains("$script:totalCount++", script, StringComparison.Ordinal);
        Assert.Contains("RecordType = 'Summary'", script, StringComparison.Ordinal);
        Assert.Contains("IsTotalCountExact = $isTotalCountExact", script, StringComparison.Ordinal);
        Assert.DoesNotContain("MessageTraceWindowedPagination", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildGetMessageTraceScript_RequiresV2Cmdlet()
    {
        var script = ExoMessageTraceCommands.BuildGetMessageTraceScript(new GetMessageTraceRequest
        {
            StartDate = new DateTime(2026, 3, 1, 0, 0, 0, DateTimeKind.Utc),
            EndDate = new DateTime(2026, 3, 14, 23, 59, 59, DateTimeKind.Utc),
            Page = 1,
            PageSize = 50
        });

        Assert.Contains("Get-MessageTraceV2 is required. Install/upgrade ExchangeOnlineManagement and reconnect.", script, StringComparison.Ordinal);
        Assert.DoesNotContain("Get-MessageTrace @params", script, StringComparison.Ordinal);
        Assert.DoesNotContain("LegacyMessageTraceCmdlet", script, StringComparison.Ordinal);
        Assert.DoesNotContain("$legacyPage++", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildGetMobileDeviceStatisticsScript_UsesSingleIdentity()
    {
        var script = ExoMobileDeviceCommands.BuildGetMobileDeviceStatisticsScript("device-01");

        Assert.Contains("Get-MobileDeviceStatistics -Identity 'device-01' -ErrorAction Stop", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildGetMobileDeviceCasMailboxScript_UsesSingleMailboxIdentity()
    {
        var script = ExoMobileDeviceCommands.BuildGetMobileDeviceCasMailboxScript("user@contoso.com");

        Assert.Contains("Get-CASMailbox -Identity 'user@contoso.com' -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("Resolve-UserPrincipalName $cas 'user@contoso.com'", script, StringComparison.Ordinal);
        Assert.Contains("UserPrincipalName = $userPrincipalName", script, StringComparison.Ordinal);
        Assert.Contains("MailboxIdentity = $resolvedMailboxIdentity", script, StringComparison.Ordinal);
    }

    [Fact]
    public void ResolveDisplayedMailboxIdentity_PrefersUserPrincipalNameWhenAvailable()
    {
        var result = ExoMobileDeviceCommands.ResolveDisplayedMailboxIdentity(
            "8f7f9c5d-2cf2-4a8b-a385-0f814bc20469",
            "user@contoso.com");

        Assert.Equal("user@contoso.com", result);
    }

    [Fact]
    public void ResolveDisplayedMailboxIdentity_KeepsExistingMailboxIdentityWhenUpnIsMissing()
    {
        var result = ExoMobileDeviceCommands.ResolveDisplayedMailboxIdentity(
            "mailbox-guid-or-alias",
            null);

        Assert.Equal("mailbox-guid-or-alias", result);
    }

    [Fact]
    public void ResolveDisplayedMailboxLabel_PrefersFriendlyDisplayName()
    {
        var result = ExoMobileDeviceCommands.ResolveDisplayedMailboxLabel(
            "mailbox-guid-or-alias",
            "Andrea Abdel Latif - Biofer SpA",
            "user@contoso.com");

        Assert.Equal("Andrea Abdel Latif - Biofer SpA", result);
    }

    [Fact]
    public void ResolveDisplayedMailboxLabel_FallsBackToUserPrincipalNameThenIdentity()
    {
        var upnResult = ExoMobileDeviceCommands.ResolveDisplayedMailboxLabel(
            "mailbox-guid-or-alias",
            null,
            "user@contoso.com");
        var identityResult = ExoMobileDeviceCommands.ResolveDisplayedMailboxLabel(
            "mailbox-guid-or-alias",
            null,
            null);

        Assert.Equal("user@contoso.com", upnResult);
        Assert.Equal("mailbox-guid-or-alias", identityResult);
    }

    [Fact]
    public void BuildGetMailboxCasSettingsScript_UsesSingleMailboxIdentity()
    {
        var script = ExoMailboxScriptFactory.BuildGetMailboxCasSettingsScript("shared@contoso.com");

        Assert.Contains("Get-CASMailbox -Identity 'shared@contoso.com' -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("OwaEnabled = $cas.OWAEnabled", script, StringComparison.Ordinal);
        Assert.Contains("SmtpClientAuthenticationDisabled = $cas.SmtpClientAuthenticationDisabled", script, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0, 50, 51)]
    [InlineData(10, 25, 36)]
    public void BuildGetMailboxesScript_UsesWindowedPreloadWithoutUnlimitedWhenSearchIsEmpty(int skip, int pageSize, int expectedWindowSize)
    {
        var script = ExoMailboxCommands.BuildGetMailboxesScript(
            skip,
            pageSize,
            filterParam: "-Filter \"RecipientTypeDetails -eq 'UserMailbox'\"",
            escapedSearch: string.Empty,
            sortProperty: "DisplayName",
            sortDirection: string.Empty,
            useWindowedLoad: true);

        Assert.Contains($"$pageWindowSize = {expectedWindowSize}", script, StringComparison.Ordinal);
        Assert.Contains("Get-Mailbox -ResultSize $pageWindowSize", script, StringComparison.Ordinal);
        Assert.DoesNotContain("Get-Mailbox -ResultSize Unlimited", script, StringComparison.Ordinal);
        Assert.Contains("IsTotalCountExact = $false", script, StringComparison.Ordinal);
        Assert.Contains("HasMore = $hasMore", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildGetMailboxesScript_KeepsExactFullScanWhenSearchIsPresent()
    {
        var script = ExoMailboxCommands.BuildGetMailboxesScript(
            skip: 0,
            pageSize: 50,
            filterParam: string.Empty,
            escapedSearch: "mario",
            sortProperty: "DisplayName",
            sortDirection: string.Empty,
            useWindowedLoad: false);

        Assert.Contains("Get-Mailbox -ResultSize Unlimited", script, StringComparison.Ordinal);
        Assert.Contains("DisplayName -like '*mario*'", script, StringComparison.Ordinal);
        Assert.Contains("IsTotalCountExact = $true", script, StringComparison.Ordinal);
    }

}

