using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Worker.PowerShell;

namespace OnlyExo365.Tests;

public sealed class AuditPowerShellRegressionTests
{
    [Theory]
    [InlineData("Mailbox")]
    [InlineData("MailUser")]
    [InlineData("MigrationCreate")]
    [InlineData("MigrationTest")]
    public async Task GeneratedCredentialCommands_BindNamedParameters(string kind)
    {
        using var fixture = new LocalPowerShellFixture();
        // Local cmdlets validate binding without contacting Microsoft 365.
        await fixture.DefineCommandsAsync("""
            function New-Mailbox { param($Password) if ($null -eq $Password) { throw 'Password not bound' }; 'bound' }
            function New-MailUser { param($Password) if ($null -eq $Password) { throw 'Password not bound' }; 'bound' }
            function New-MigrationEndpoint { param($Credentials) if ($null -eq $Credentials -or $Credentials.UserName -ne 'dummy-user') { throw 'Credentials not bound' } }
            function Test-MigrationServerAvailability { param($Credentials) if ($null -eq $Credentials -or $Credentials.UserName -ne 'dummy-user') { throw 'Credentials not bound' }; [pscustomobject]@{ Result = 'bound' } }
            """);

        var command = kind switch
        {
            "Mailbox" => ExoMailboxScriptFactory.BuildCreateMailboxCommand(new CreateMailboxRequest
            {
                DisplayName = "Dummy", Alias = "dummy", PrimarySmtpAddress = "dummy@example.com", MailboxType = "User", Password = "dummy-password"
            }),
            "MailUser" => ExoCommands.BuildUpsertContactCommand(new UpsertContactRequest
            {
                ContactKind = "MailUser", Name = "Dummy", DisplayName = "Dummy", Alias = "dummy", PrimarySmtpAddress = "dummy@example.com",
                ExternalEmailAddress = "dummy@example.org", UserPrincipalName = "dummy@example.com", Password = "dummy-password"
            }),
            "MigrationCreate" => MigrationCommandBuilder.BuildUpsertMigrationEndpointCommand(new UpsertMigrationEndpointRequest
            {
                Name = "Dummy", EndpointType = "ExchangeRemoteMove", RemoteServer = "example.com", Username = "dummy-user", Password = "dummy-password"
            }),
            _ => MigrationCommandBuilder.BuildTestMigrationEndpointCommand(new TestMigrationEndpointRequest
            {
                EndpointType = "ExchangeRemoteMove", RemoteServer = "example.com", Username = "dummy-user", Password = "dummy-password"
            })
        };

        var result = await fixture.Engine.ExecuteAsync(command.Script, command.Parameters);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEmpty(result.Output);
        Assert.DoesNotContain("dummy-password", command.Script, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ErrorContainingDeprecated_RemainsFailure()
    {
        using var fixture = new LocalPowerShellFixture();
        var result = await fixture.Engine.ExecuteAsync("Write-Error 'This endpoint is deprecated and unavailable'");
        Assert.False(result.Success);
        Assert.Single(result.Errors);
        Assert.Contains("deprecated", result.ErrorMessage, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("Mailboxes")]
    [InlineData("DeletedMailboxes")]
    [InlineData("Contacts")]
    [InlineData("Resources")]
    [InlineData("PublicFolders")]
    [InlineData("Migrations")]
    [InlineData("MobileDevices")]
    [InlineData("GroupMembers")]
    public async Task ListingFailure_DoesNotBecomeSuccessfulEmptyList(string area)
    {
        using var fixture = new LocalPowerShellFixture();
        await fixture.DefineCommandsAsync("""
            function Get-Mailbox { throw 'Access denied by local test cmdlet' }
            function Get-EXOMailbox { throw 'Access denied by local test cmdlet' }
            function Get-MailContact { throw 'Access denied by local test cmdlet' }
            function Get-MailUser { throw 'Access denied by local test cmdlet' }
            function Get-PublicFolder { throw 'Access denied by local test cmdlet' }
            function Get-MigrationBatch { throw 'Access denied by local test cmdlet' }
            function Get-MobileDevice { throw 'Access denied by local test cmdlet' }
            function Get-CASMailbox { throw 'Access denied by local test cmdlet' }
            function Get-DistributionGroupMember { throw 'Access denied by local test cmdlet' }
            """);
        var engine = fixture.Engine;
        Func<Task> action = area switch
        {
            "Mailboxes" => async () => await new ExoMailboxListingCommands(engine).GetMailboxesAsync(new GetMailboxesRequest()),
            "DeletedMailboxes" => async () => await new ExoMailboxListingCommands(engine).GetDeletedMailboxesAsync(new GetDeletedMailboxesRequest()),
            "Contacts" => async () => await new ExoContactCommands(engine).GetContactsAsync(new GetContactsRequest()),
            "Resources" => async () => await new ExoResourceCommands(engine, new ExoMailboxReportingCommands(engine)).GetResourceMailboxesAsync(new GetResourceMailboxesRequest()),
            "PublicFolders" => async () => await new ExoPublicFolderCommands(engine).GetPublicFoldersAsync(new GetPublicFoldersRequest()),
            "MobileDevices" => async () => await new ExoMobileDeviceCommands(engine, new CapabilityDetector(engine)).GetMobileDevicesAsync(new GetMobileDevicesRequest()),
            "GroupMembers" => async () => await new ExoGroupCommands(engine, new CapabilityDetector(engine)).GetGroupMembersPageAsync("dummy", "Distribution", 0, 10, null, default),
            _ => async () => await new ExoMigrationCommands(engine).GetMigrationBatchesAsync(new GetMigrationBatchesRequest())
        };
        await Assert.ThrowsAsync<InvalidOperationException>(action);
    }
}
