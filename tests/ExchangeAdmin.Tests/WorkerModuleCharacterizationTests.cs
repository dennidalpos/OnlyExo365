namespace ExchangeAdmin.Tests;

public sealed class WorkerModuleCharacterizationTests
{
    [Fact]
    public void ExoDashboardCommands_TracksFallbackAndPartialDataWarnings()
    {
        var content = ReadWorker("ExoDashboardCommands.cs");

        Assert.Contains("AddWarning(", content, StringComparison.Ordinal);
        Assert.Contains("MailboxCountFallbackUsed", content, StringComparison.Ordinal);
        Assert.Contains("GroupCountFallbackUsed", content, StringComparison.Ordinal);
        Assert.Contains("TenantLicensesUnavailable", content, StringComparison.Ordinal);
        Assert.Contains("AdminUsersUnavailable", content, StringComparison.Ordinal);
        Assert.Contains("stats.HasPartialData = warningDetails.Any", content, StringComparison.Ordinal);
        Assert.Contains("SampleItems = sampleItems?", content, StringComparison.Ordinal);
    }

    [Fact]
    public void ExoContactCommands_CombinesContactsAndMailUsersAndKeepsPasswordsOutOfScripts()
    {
        var content = ReadWorker("ExoContactCommands.cs");

        Assert.Contains("Get-MailContact -ResultSize Unlimited", content, StringComparison.Ordinal);
        Assert.Contains("Get-MailUser -ResultSize Unlimited", content, StringComparison.Ordinal);
        Assert.Contains("Unsupported ContactKind ignored", content, StringComparison.Ordinal);
        Assert.Contains("Unsupported SortBy ignored", content, StringComparison.Ordinal);
        Assert.Contains("ConvertTo-SecureString $PlainTextPassword -AsPlainText -Force", content, StringComparison.Ordinal);
        Assert.Contains("[\"PlainTextPassword\"] = request.Password ?? string.Empty", content, StringComparison.Ordinal);
    }

    [Fact]
    public void ExoMailSecurityCommands_LoadsEachAreaOnlyWhenSupported()
    {
        var content = ReadWorker("ExoMailSecurityCommands.cs");

        Assert.Contains("LoadIfSupportedAsync(", content, StringComparison.Ordinal);
        Assert.Contains("Get-DkimSigningConfig", content, StringComparison.Ordinal);
        Assert.Contains("Get-HostedContentFilterPolicy", content, StringComparison.Ordinal);
        Assert.Contains("Get-AntiPhishPolicy", content, StringComparison.Ordinal);
        Assert.Contains("Get-MalwareFilterPolicy", content, StringComparison.Ordinal);
        Assert.Contains("Get-QuarantinePolicy", content, StringComparison.Ordinal);
        Assert.Contains("Get-HostedOutboundSpamFilterPolicy", content, StringComparison.Ordinal);
        Assert.Contains("warnings.Add($\"{areaName}: cmdlet {commandName} is not available in the current session.\")", content, StringComparison.Ordinal);
    }

    [Fact]
    public void ExoMigrationCommands_ValidatesPreflightAndSupportsEndpointAndBatchLifecycle()
    {
        var content = ReadWorker("ExoMigrationCommands.cs");

        Assert.Contains("Unsupported migration batch status ignored", content, StringComparison.Ordinal);
        Assert.Contains("Get-MigrationBatch -ResultSize Unlimited", content, StringComparison.Ordinal);
        Assert.Contains("Get-MigrationEndpoint -ErrorAction Stop", content, StringComparison.Ordinal);
        Assert.Contains("CSV file not found.", content, StringComparison.Ordinal);
        Assert.Contains("A migration batch with the same name already exists.", content, StringComparison.Ordinal);
        Assert.Contains("Start-MigrationBatch -Identity", content, StringComparison.Ordinal);
        Assert.Contains("Complete-MigrationBatch -Identity", content, StringComparison.Ordinal);
        Assert.Contains("Remove-MigrationBatch -Identity", content, StringComparison.Ordinal);
    }

    [Fact]
    public void ExoResourceCommands_LoadsCalendarProcessingAndPermissionDetails()
    {
        var content = ReadWorker("ExoResourceCommands.cs");

        Assert.Contains("Get-CalendarProcessing -Identity", content, StringComparison.Ordinal);
        Assert.Contains("_mailboxReportingCommands.GetMailboxPermissionsAsync", content, StringComparison.Ordinal);
        Assert.Contains("Set-CalendarProcessing -Identity", content, StringComparison.Ordinal);
        Assert.Contains("NormalizeResourceType", content, StringComparison.Ordinal);
        Assert.Contains("NormalizeAutomateProcessing", content, StringComparison.Ordinal);
    }

    [Fact]
    public void ExoMailboxDetailCommands_ConditionallyEnrichesMailboxDetails()
    {
        var content = ReadWorker("ExoMailboxDetailCommands.cs");

        Assert.Contains("if (request.IncludeStatistics)", content, StringComparison.Ordinal);
        Assert.Contains("if (request.IncludeRules)", content, StringComparison.Ordinal);
        Assert.Contains("if (request.IncludeAutoReply)", content, StringComparison.Ordinal);
        Assert.Contains("if (request.IncludePermissions)", content, StringComparison.Ordinal);
        Assert.Contains("await TryEnrichCasMailboxSettingsAsync", content, StringComparison.Ordinal);
        Assert.Contains("_mailboxLicenseCommands.GetUserLicensesAsync", content, StringComparison.Ordinal);
    }

    [Fact]
    public void ExoMailboxLifecycleCommands_UsesDedicatedCreateRestoreAndConvertFlows()
    {
        var content = ReadWorker("ExoMailboxLifecycleCommands.cs");

        Assert.Contains("BuildCreateMailboxCommand", content, StringComparison.Ordinal);
        Assert.Contains("ConvertMailboxTypeAsync(request.Identity, \"Shared\", \"shared\"", content, StringComparison.Ordinal);
        Assert.Contains("ConvertMailboxTypeAsync(request.Identity, \"Regular\", \"regular\"", content, StringComparison.Ordinal);
        Assert.Contains("BuildRestoreMailboxCommand", content, StringComparison.Ordinal);
        Assert.Contains("ExoMailboxMapper.ToRestoreMailboxResponse", content, StringComparison.Ordinal);
    }

    [Fact]
    public void ExchangeConnectionSession_ResetsGraphStateAndBootstrapsGraphAndCompliance()
    {
        var content = ReadWorker("ExchangeConnectionSession.cs");

        Assert.Contains("Disconnect-MgGraph -ErrorAction SilentlyContinue", content, StringComparison.Ordinal);
        Assert.Contains("_engine.GraphConnected = false;", content, StringComparison.Ordinal);
        Assert.Contains("await ConnectInitialServiceBundleAsync", content, StringComparison.Ordinal);
        Assert.Contains("ConnectMicrosoftGraphAsync(", content, StringComparison.Ordinal);
        Assert.Contains("ConnectComplianceAsync(", content, StringComparison.Ordinal);
        Assert.Contains("HasRequiredScopes", content, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerShellExecutionPipeline_HandlesRecoveryCancellationAndRunspaceCorruptionSignals()
    {
        var content = ReadWorker("PowerShellExecutionPipeline.cs");

        Assert.Contains("_runspaceRecoveryService.IsRunspaceUsable()", content, StringComparison.Ordinal);
        Assert.Contains("TryRecoverRunspaceAsync", content, StringComparison.Ordinal);
        Assert.Contains("cancellationToken.Register(() =>", content, StringComparison.Ordinal);
        Assert.Contains("ps.Stop();", content, StringComparison.Ordinal);
        Assert.Contains("RunspaceCorrupted = true", content, StringComparison.Ordinal);
        Assert.Contains("errorMessage.Contains(\"deprecat\"", content, StringComparison.Ordinal);
    }

    [Fact]
    public void RunspaceLifecycleManager_UsesRemoteSignedAndBestEffortPackageManagementPreparation()
    {
        var content = ReadWorker("RunspaceLifecycleManager.cs");

        Assert.Contains("iss.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.RemoteSigned;", content, StringComparison.Ordinal);
        Assert.Contains("EnsurePackageManagementAvailableAsync", content, StringComparison.Ordinal);
        Assert.Contains("[Version]'1.4.4'", content, StringComparison.Ordinal);
        Assert.Contains("Import-Module -Name $module.Path -Global -ErrorAction Stop | Out-Null", content, StringComparison.Ordinal);
        Assert.Contains("ImportModuleAsync", content, StringComparison.Ordinal);
    }

    [Fact]
    public void RunspaceRecoveryService_RecreatesRunspaceAndResetsConnectionFlags()
    {
        var content = ReadWorker("RunspaceRecoveryService.cs");

        Assert.Contains("_runspaceLifecycleManager.RecreateRunspaceAsync(_engine.ModuleAvailable)", content, StringComparison.Ordinal);
        Assert.Contains("_engine.ConsecutiveFailures = 0;", content, StringComparison.Ordinal);
        Assert.Contains("_engine.Connected = false;", content, StringComparison.Ordinal);
        Assert.Contains("_engine.GraphConnected = false;", content, StringComparison.Ordinal);
    }

    private static string ReadWorker(string fileName)
        => File.ReadAllText(TestPathHelper.GetRepositoryPath("src", "ExchangeAdmin.Worker", "PowerShell", fileName));
}
