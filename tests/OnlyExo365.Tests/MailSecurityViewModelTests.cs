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

    [Fact]
    public void Draft_CancelledPolicySelectionAndRefreshPreserveEditedAction()
    {
        var worker = new MailSecurityTestWorkerService();
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var editor = new MailSecurityViewModel(worker, shell);
        var first = new HostedContentFilterPolicyDto { Identity = "a", Name = "A" };
        editor.SelectedAntiSpamPolicy = first;
        Assert.False(editor.HasPendingChanges);
        editor.AntiSpamSpamAction = "CustomAction";
        Assert.True(editor.HasPendingChanges);
        ErrorDialogService.ConfirmationHandlerOverride = (_, _) => false;
        try
        {
            editor.SelectedAntiSpamPolicy = new HostedContentFilterPolicyDto { Identity = "b", Name = "B" };
            editor.RefreshWorkspaceCommand.Execute(null);
            Assert.Same(first, editor.SelectedAntiSpamPolicy);
            Assert.Equal("CustomAction", editor.AntiSpamSpamAction);
            Assert.Equal(0, worker.GetMailSecurityBaselineCalls);
        }
        finally { ErrorDialogService.ConfirmationHandlerOverride = null; editor.ReleaseUiResources(); }
    }

    [Fact]
    public void SaveAntiSpam_SubmitsSnapshotRatherThanValuesChangedDuringConfirmation()
    {
        var worker = new MailSecurityTestWorkerService();
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var editor = new MailSecurityViewModel(worker, shell)
        {
            SelectedAntiSpamPolicy = new HostedContentFilterPolicyDto { Identity = "a", Name = "A", SpamAction = "Delete" }
        };
        editor.AntiSpamSpamAction = "Quarantine";
        ErrorDialogService.ConfirmationHandlerOverride = (_, message) =>
        {
            Assert.Contains("A", message, StringComparison.Ordinal);
            editor.AntiSpamSpamAction = "Redirect";
            return true;
        };
        try
        {
            editor.SaveAntiSpamCommand.Execute(null);
            Assert.Equal("a", worker.LastAntiSpamRequest?.Identity);
            Assert.Equal("Quarantine", worker.LastAntiSpamRequest?.SpamAction);
        }
        finally { ErrorDialogService.ConfirmationHandlerOverride = null; editor.ReleaseUiResources(); }
    }

    [Fact]
    public async Task SaveAntiSpam_CancelledDiscardOfOtherDraftsPreservesEditorsAndSkipsWrite()
    {
        var worker = new MailSecurityTestWorkerService();
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var editor = new MailSecurityViewModel(worker, shell);
        await editor.LoadAsync();
        var selected = editor.SelectedAntiSpamPolicy;
        editor.AntiSpamSpamAction = "Delete";
        editor.AntiPhishThresholdLevel = 4;
        var prompts = 0;
        ErrorDialogService.ConfirmationHandlerOverride = (_, message) =>
        {
            prompts++;
            Assert.Contains(OnlyExo365.Shell.Localization.Loc.Get("Security.SaveDiscardsOtherDrafts"), message, StringComparison.Ordinal);
            return false;
        };
        try
        {
            editor.SaveAntiSpamCommand.Execute(null);
            Assert.Equal(1, prompts);
            Assert.Null(worker.LastAntiSpamRequest);
            Assert.Equal(1, worker.GetMailSecurityBaselineCalls);
            Assert.Same(selected, editor.SelectedAntiSpamPolicy);
            Assert.Equal("Delete", editor.AntiSpamSpamAction);
            Assert.Equal(4, editor.AntiPhishThresholdLevel);
            Assert.True(editor.HasPendingChanges);
        }
        finally { ErrorDialogService.ConfirmationHandlerOverride = null; editor.ReleaseUiResources(); }
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
        public UpdateHostedContentFilterPolicyRequest? LastAntiSpamRequest { get; private set; }

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
        {
            LastAntiSpamRequest = request;
            return Task.FromResult(Result.Success());
        }

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

