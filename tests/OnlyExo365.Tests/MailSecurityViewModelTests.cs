using System.Reflection;
using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;
using OnlyExo365.Shell.Ipc;
using OnlyExo365.Shell.ViewModels;

namespace OnlyExo365.Tests;

public sealed class MailSecurityViewModelTests
{
    [Fact]
    public async Task LoadAsync_PopulatesMailSecurityCollectionsAndWarnings()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var worker = new MailSecurityTestWorkerService();
        var viewModel = new MailSecurityViewModel(worker, shell);

        await viewModel.LoadAsync();

        Assert.Single(viewModel.DkimConfigs);
        Assert.Single(viewModel.AntiSpamPolicies);
        Assert.Single(viewModel.AntiPhishPolicies);
        Assert.Single(viewModel.MalwarePolicies);
        Assert.Single(viewModel.QuarantinePolicies);
        Assert.Single(viewModel.OutboundSpamPolicies);
        Assert.True(viewModel.HasWarnings);
        Assert.Contains("cmdlet", viewModel.WarningsText, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(1, worker.GetMailSecurityBaselineCalls);
    }

    [Fact]
    public void SelectingAntiSpamPolicy_PrefillsEditorFromSelection()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var viewModel = new MailSecurityViewModel(new MailSecurityTestWorkerService(), shell);

        viewModel.SelectedAntiSpamPolicy = new HostedContentFilterPolicyDto
        {
            Identity = "policy-01",
            Name = "Preset",
            RuleIdentity = "rule-01",
            RuleState = "Enabled",
            BulkThreshold = 7,
            SpamAction = "MoveToJmf",
            HighConfidenceSpamAction = "Quarantine",
            PhishSpamAction = "Delete"
        };

        Assert.True(viewModel.SelectedAntiSpamHasRule);
        Assert.True(viewModel.AntiSpamRuleEnabled);
        Assert.Equal(7, viewModel.AntiSpamBulkThreshold);
        Assert.Equal("Delete", viewModel.AntiSpamPhishSpamAction);
    }

    private static void SetExchangeConnected(ShellViewModel shell)
    {
        var property = typeof(ShellViewModel)
            .GetProperty(nameof(ShellViewModel.ExchangeState), BindingFlags.Instance | BindingFlags.Public);

        Assert.NotNull(property);
        property!.SetValue(shell, ConnectionState.Connected);
    }

    private sealed class MailSecurityTestWorkerService : TestMailSecurityWorkerServiceBase
    {
        public int GetMailSecurityBaselineCalls { get; private set; }

        public override Task<Result<GetMailSecurityBaselineResponse>> GetMailSecurityBaselineAsync(GetMailSecurityBaselineRequest? request = null, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            GetMailSecurityBaselineCalls++;
            return Task.FromResult(Result<GetMailSecurityBaselineResponse>.Success(new GetMailSecurityBaselineResponse
            {
                DkimConfigs = [new DkimSigningConfigDto { Identity = "contoso.com", Domain = "contoso.com", Enabled = true }],
                AntiSpamPolicies = [new HostedContentFilterPolicyDto { Identity = "spam-01", Name = "Spam", RuleIdentity = "rule-spam", RuleState = "Enabled", BulkThreshold = 6 }],
                AntiPhishPolicies = [new AntiPhishPolicyDto { Identity = "phish-01", Name = "Phish", RuleIdentity = "rule-phish", RuleState = "Enabled", PhishThresholdLevel = 2 }],
                MalwarePolicies = [new MalwareFilterPolicyDto { Identity = "mal-01", Name = "Malware", RuleIdentity = "rule-mal", RuleState = "Disabled", ZapEnabled = true }],
                QuarantinePolicies = [new QuarantinePolicyDto { Identity = "quar-01", Name = "Quarantine", EndUserQuarantinePermissionsValue = "Preview" }],
                OutboundSpamPolicies = [new HostedOutboundSpamFilterPolicyDto { Identity = "out-01", Name = "Outbound", RecipientLimitPerDay = 1000 }],
            Warnings = ["Quarantine: cmdlet Set-QuarantinePolicy is not available in the current session."]
            }));
        }

        public override Task<Result> UpdateDkimSigningConfigAsync(UpdateDkimSigningConfigRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => throw TestStubExceptions.CreateUnsupported();

        public override Task<Result> UpdateHostedContentFilterPolicyAsync(UpdateHostedContentFilterPolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => throw TestStubExceptions.CreateUnsupported();

        public override Task<Result> UpdateAntiPhishPolicyAsync(UpdateAntiPhishPolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => throw TestStubExceptions.CreateUnsupported();

        public override Task<Result> UpdateMalwareFilterPolicyAsync(UpdateMalwareFilterPolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => throw TestStubExceptions.CreateUnsupported();

        public override Task<Result> UpdateQuarantinePolicyAsync(UpdateQuarantinePolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => throw TestStubExceptions.CreateUnsupported();

        public override Task<Result> UpdateHostedOutboundSpamFilterPolicyAsync(UpdateHostedOutboundSpamFilterPolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => throw TestStubExceptions.CreateUnsupported();
    }
}

