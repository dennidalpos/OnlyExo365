using System.Collections.ObjectModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

internal sealed class MailFlowRulesViewModel : MailFlowSectionViewModelBase
{
    private static readonly string[] ValidationProperties =
    {
        nameof(IsRuleInputValid),
        nameof(HasAnyRuleCondition),
        nameof(HasAnyRuleAction),
        nameof(IsRuleTestInputValid),
        nameof(RuleValidationMessage),
        nameof(TestValidationMessage)
    };

    private TransportRuleDto? _selectedRule;
    private string? _ruleIdentity;
    private string _ruleName = string.Empty;
    private string _ruleFrom = string.Empty;
    private string _ruleSentTo = string.Empty;
    private string _ruleSenderDomainIs = string.Empty;
    private string _ruleRecipientDomainIs = string.Empty;
    private string _ruleSentToMemberOf = string.Empty;
    private string _ruleSubjectContains = string.Empty;
    private string _ruleExceptIfFrom = string.Empty;
    private string _ruleExceptIfSentTo = string.Empty;
    private string _ruleExceptIfSenderDomainIs = string.Empty;
    private string _ruleExceptIfRecipientDomainIs = string.Empty;
    private string _ruleExceptIfSubjectContains = string.Empty;
    private string _rulePrependSubject = string.Empty;
    private string _ruleRedirectMessageTo = string.Empty;
    private string _ruleBlindCopyTo = string.Empty;
    private string _ruleAddToRecipients = string.Empty;
    private bool _ruleStopRuleProcessing;
    private bool _ruleDeleteMessage;
    private string _ruleMode = "Enforce";
    private bool _ruleEnabled = true;
    private string _testSender = string.Empty;
    private string _testRecipient = string.Empty;
    private string _testSubject = string.Empty;
    private string _testResult = string.Empty;

    public MailFlowRulesViewModel(
        IMailFlowWorkerService workerService,
        ShellViewModel shellViewModel,
        MailFlowOperationCoordinator coordinator,
        Func<CancellationToken, Task> refreshAllAsync)
        : base(workerService, shellViewModel, coordinator, refreshAllAsync)
    {
        NewRuleCommand = new AsyncRelayCommand(NewRuleAsync, () => !Coordinator.IsLoading);
        EnableRuleCommand = new AsyncRelayCommand(ct => SetRuleStateAsync(true, ct), () => SelectedRule != null && !Coordinator.IsLoading);
        DisableRuleCommand = new AsyncRelayCommand(ct => SetRuleStateAsync(false, ct), () => SelectedRule != null && !Coordinator.IsLoading);
        SaveRuleCommand = new AsyncRelayCommand(SaveRuleAsync, () => !Coordinator.IsLoading && IsRuleInputValid);
        RemoveRuleCommand = new AsyncRelayCommand(RemoveRuleAsync, () => !Coordinator.IsLoading && SelectedRule != null);
        TestRuleCommand = new AsyncRelayCommand(TestRuleAsync, () => !Coordinator.IsLoading && IsRuleTestInputValid);
    }

    public IReadOnlyList<string> RuleModes { get; } = new[] { "Enforce", "Audit" };
    public ObservableCollection<TransportRuleDto> TransportRules { get; } = new();

    public TransportRuleDto? SelectedRule
    {
        get => _selectedRule;
        set
        {
            if (SetProperty(ref _selectedRule, value))
            {
                if (value != null)
                {
                    RuleIdentity = value.Identity;
                    RuleName = value.Name;
                    RuleFrom = string.Join(",", value.From);
                    RuleSentTo = string.Join(",", value.SentTo);
                    RuleSenderDomainIs = string.Join(",", value.SenderDomainIs);
                    RuleRecipientDomainIs = string.Join(",", value.RecipientDomainIs);
                    RuleSentToMemberOf = string.Join(",", value.SentToMemberOf);
                    RuleSubjectContains = string.Join(",", value.SubjectContainsWords);
                    RuleExceptIfFrom = string.Join(",", value.ExceptIfFrom);
                    RuleExceptIfSentTo = string.Join(",", value.ExceptIfSentTo);
                    RuleExceptIfSenderDomainIs = string.Join(",", value.ExceptIfSenderDomainIs);
                    RuleExceptIfRecipientDomainIs = string.Join(",", value.ExceptIfRecipientDomainIs);
                    RuleExceptIfSubjectContains = string.Join(",", value.ExceptIfSubjectContainsWords);
                    RulePrependSubject = value.PrependSubject;
                    RuleRedirectMessageTo = string.Join(",", value.RedirectMessageTo);
                    RuleBlindCopyTo = string.Join(",", value.BlindCopyTo);
                    RuleAddToRecipients = string.Join(",", value.AddToRecipients);
                    RuleStopRuleProcessing = value.StopRuleProcessing;
                    RuleDeleteMessage = value.DeleteMessage;
                    RuleMode = string.IsNullOrWhiteSpace(value.Mode) ? "Enforce" : value.Mode;
                    RuleEnabled = !string.Equals(value.State, "Disabled", StringComparison.OrdinalIgnoreCase);
                }
                else
                {
                    ResetEditor();
                }

                OnPropertyChanged(nameof(CanEditSelectedRule));
                InvalidateCommands();
            }
        }
    }

    public bool CanEditSelectedRule => SelectedRule != null && !Coordinator.IsLoading;
    public bool HasAnyRuleCondition =>
        MailFlowViewModelSupport.SplitCsv(RuleFrom).Count > 0 ||
        MailFlowViewModelSupport.SplitCsv(RuleSentTo).Count > 0 ||
        MailFlowViewModelSupport.SplitCsv(RuleSenderDomainIs).Count > 0 ||
        MailFlowViewModelSupport.SplitCsv(RuleRecipientDomainIs).Count > 0 ||
        MailFlowViewModelSupport.SplitCsv(RuleSentToMemberOf).Count > 0 ||
        MailFlowViewModelSupport.SplitCsv(RuleSubjectContains).Count > 0;

    public bool HasAnyRuleAction =>
        !string.IsNullOrWhiteSpace(RulePrependSubject) ||
        MailFlowViewModelSupport.SplitCsv(RuleRedirectMessageTo).Count > 0 ||
        MailFlowViewModelSupport.SplitCsv(RuleBlindCopyTo).Count > 0 ||
        MailFlowViewModelSupport.SplitCsv(RuleAddToRecipients).Count > 0 ||
        RuleStopRuleProcessing ||
        RuleDeleteMessage;

    public bool IsRuleInputValid => !string.IsNullOrWhiteSpace(RuleName) && RuleModes.Contains(RuleMode) && HasAnyRuleCondition && HasAnyRuleAction;
    public bool IsRuleTestInputValid => MailFlowViewModelSupport.IsValidEmail(TestSender) && MailFlowViewModelSupport.IsValidEmail(TestRecipient);
    public string RuleValidationMessage => IsRuleInputValid ? string.Empty : "Rule: name is required, mode must be valid, and at least one condition and one action are required.";
    public string TestValidationMessage => IsRuleTestInputValid ? string.Empty : "Test: sender and recipient must be valid email addresses.";

    public string? RuleIdentity { get => _ruleIdentity; set => SetProperty(ref _ruleIdentity, value); }
    public string RuleName { get => _ruleName; set => SetEditorProperty(ref _ruleName, value); }
    public string RuleFrom { get => _ruleFrom; set => SetEditorProperty(ref _ruleFrom, value); }
    public string RuleSentTo { get => _ruleSentTo; set => SetEditorProperty(ref _ruleSentTo, value); }
    public string RuleSenderDomainIs { get => _ruleSenderDomainIs; set => SetEditorProperty(ref _ruleSenderDomainIs, value); }
    public string RuleRecipientDomainIs { get => _ruleRecipientDomainIs; set => SetEditorProperty(ref _ruleRecipientDomainIs, value); }
    public string RuleSentToMemberOf { get => _ruleSentToMemberOf; set => SetEditorProperty(ref _ruleSentToMemberOf, value); }
    public string RuleSubjectContains { get => _ruleSubjectContains; set => SetEditorProperty(ref _ruleSubjectContains, value); }
    public string RuleExceptIfFrom { get => _ruleExceptIfFrom; set => SetEditorProperty(ref _ruleExceptIfFrom, value); }
    public string RuleExceptIfSentTo { get => _ruleExceptIfSentTo; set => SetEditorProperty(ref _ruleExceptIfSentTo, value); }
    public string RuleExceptIfSenderDomainIs { get => _ruleExceptIfSenderDomainIs; set => SetEditorProperty(ref _ruleExceptIfSenderDomainIs, value); }
    public string RuleExceptIfRecipientDomainIs { get => _ruleExceptIfRecipientDomainIs; set => SetEditorProperty(ref _ruleExceptIfRecipientDomainIs, value); }
    public string RuleExceptIfSubjectContains { get => _ruleExceptIfSubjectContains; set => SetEditorProperty(ref _ruleExceptIfSubjectContains, value); }
    public string RulePrependSubject { get => _rulePrependSubject; set => SetEditorProperty(ref _rulePrependSubject, value); }
    public string RuleRedirectMessageTo { get => _ruleRedirectMessageTo; set => SetEditorProperty(ref _ruleRedirectMessageTo, value); }
    public string RuleBlindCopyTo { get => _ruleBlindCopyTo; set => SetEditorProperty(ref _ruleBlindCopyTo, value); }
    public string RuleAddToRecipients { get => _ruleAddToRecipients; set => SetEditorProperty(ref _ruleAddToRecipients, value); }
    public bool RuleStopRuleProcessing { get => _ruleStopRuleProcessing; set => SetEditorProperty(ref _ruleStopRuleProcessing, value); }
    public bool RuleDeleteMessage { get => _ruleDeleteMessage; set => SetEditorProperty(ref _ruleDeleteMessage, value); }
    public string RuleMode { get => _ruleMode; set => SetEditorProperty(ref _ruleMode, value); }
    public bool RuleEnabled { get => _ruleEnabled; set => SetProperty(ref _ruleEnabled, value); }
    public string TestSender { get => _testSender; set => SetEditorProperty(ref _testSender, value); }
    public string TestRecipient { get => _testRecipient; set => SetEditorProperty(ref _testRecipient, value); }
    public string TestSubject { get => _testSubject; set => SetProperty(ref _testSubject, value); }
    public string TestResult { get => _testResult; set => SetProperty(ref _testResult, value); }

    public ICommand NewRuleCommand { get; }
    public ICommand EnableRuleCommand { get; }
    public ICommand DisableRuleCommand { get; }
    public ICommand SaveRuleCommand { get; }
    public ICommand RemoveRuleCommand { get; }
    public ICommand TestRuleCommand { get; }

    public async Task LoadAsync(CancellationToken cancellationToken)
    {
        var result = await WorkerService.GetTransportRulesAsync(new GetTransportRulesRequest(), cancellationToken: cancellationToken);
        if (!result.IsSuccess)
        {
            var error = result.Error?.Message ?? "Unable to load transport rules.";
            SetError(error);
            ShellViewModel.AddLog(LogLevel.Error, $"MailFlow rules load failed: {error}", "MailFlow");
            return;
        }

        TransportRules.Clear();
        foreach (var item in result.Value?.Rules ?? new List<TransportRuleDto>())
        {
            TransportRules.Add(item);
        }
    }

    private Task NewRuleAsync(CancellationToken cancellationToken)
    {
        SelectedRule = null;
        ResetEditor();
        SetError(null);
        TestResult = string.Empty;
        return Task.CompletedTask;
    }

    private async Task SetRuleStateAsync(bool enabled, CancellationToken cancellationToken)
    {
        if (SelectedRule == null)
        {
            return;
        }

        if (!enabled &&
            !ConfirmMutation(
                "Disable transport rule",
                SelectedRule.Name,
                "Can stop tenant-wide controls or enforcement.",
                "Confirm tenant-wide disable"))
        {
            return;
        }

        await ExecuteBusyActionAsync(async ct =>
        {
            var result = await WorkerService.SetTransportRuleStateAsync(new SetTransportRuleStateRequest
            {
                Identity = SelectedRule.Identity,
                Enabled = enabled
            }, cancellationToken: ct);

            if (!result.IsSuccess)
            {
                var error = result.Error?.Message ?? "Rule state error";
                SetError(error);
                ShellViewModel.AddLog(LogLevel.Error, $"Set rule state failed (rule={SelectedRule.Name}): {error}", "MailFlow");
                return;
            }

            await RefreshAllAsync(ct);
        }, cancellationToken);
    }

    private async Task SaveRuleAsync(CancellationToken cancellationToken)
    {
        if (!IsRuleInputValid)
        {
            return;
        }

        await ExecuteBusyActionAsync(async ct =>
        {
            var result = await WorkerService.UpsertTransportRuleAsync(new UpsertTransportRuleRequest
            {
                Identity = string.IsNullOrWhiteSpace(RuleIdentity) ? null : RuleIdentity,
                Name = RuleName.Trim(),
                From = MailFlowViewModelSupport.SplitCsv(RuleFrom),
                SentTo = MailFlowViewModelSupport.SplitCsv(RuleSentTo),
                SenderDomainIs = MailFlowViewModelSupport.SplitCsv(RuleSenderDomainIs),
                RecipientDomainIs = MailFlowViewModelSupport.SplitCsv(RuleRecipientDomainIs),
                SentToMemberOf = MailFlowViewModelSupport.SplitCsv(RuleSentToMemberOf),
                SubjectContainsWords = MailFlowViewModelSupport.SplitCsv(RuleSubjectContains),
                ExceptIfFrom = MailFlowViewModelSupport.SplitCsv(RuleExceptIfFrom),
                ExceptIfSentTo = MailFlowViewModelSupport.SplitCsv(RuleExceptIfSentTo),
                ExceptIfSenderDomainIs = MailFlowViewModelSupport.SplitCsv(RuleExceptIfSenderDomainIs),
                ExceptIfRecipientDomainIs = MailFlowViewModelSupport.SplitCsv(RuleExceptIfRecipientDomainIs),
                ExceptIfSubjectContainsWords = MailFlowViewModelSupport.SplitCsv(RuleExceptIfSubjectContains),
                PrependSubject = string.IsNullOrWhiteSpace(RulePrependSubject) ? null : RulePrependSubject.Trim(),
                RedirectMessageTo = MailFlowViewModelSupport.SplitCsv(RuleRedirectMessageTo),
                BlindCopyTo = MailFlowViewModelSupport.SplitCsv(RuleBlindCopyTo),
                AddToRecipients = MailFlowViewModelSupport.SplitCsv(RuleAddToRecipients),
                StopRuleProcessing = RuleStopRuleProcessing,
                DeleteMessage = RuleDeleteMessage,
                Mode = RuleMode,
                Enabled = RuleEnabled
            }, cancellationToken: ct);

            if (!result.IsSuccess)
            {
                var correlationId = Guid.NewGuid().ToString("N");
                SetError($"Saving rule failed. Try again or check the logs (ref: {correlationId}).");
                ShellViewModel.AddLog(LogLevel.Error, $"[{correlationId}] Save rule failed (name={RuleName}): {result.Error?.Message}", "MailFlow");
                return;
            }

            await RefreshAllAsync(ct);
        }, cancellationToken);
    }

    private async Task RemoveRuleAsync(CancellationToken cancellationToken)
    {
        if (SelectedRule == null)
        {
            return;
        }

        if (!ConfirmMutation(
                "Deleting transport rule",
                SelectedRule.Name,
                "Permanently removes the rule with potential tenant-wide impact.",
                "Confirm tenant-wide deletion"))
        {
            return;
        }

        await ExecuteBusyActionAsync(async ct =>
        {
            var result = await WorkerService.RemoveTransportRuleAsync(new RemoveTransportRuleRequest
            {
                Identity = SelectedRule.Identity
            }, cancellationToken: ct);

            if (!result.IsSuccess)
            {
                var error = result.Error?.Message ?? "Unable to delete rule";
                SetError(error);
                ShellViewModel.AddLog(LogLevel.Error, $"Remove rule failed (rule={SelectedRule.Name}): {error}", "MailFlow");
                return;
            }

            await RefreshAllAsync(ct);
        }, cancellationToken);
    }

    private async Task TestRuleAsync(CancellationToken cancellationToken)
    {
        if (!IsRuleTestInputValid)
        {
            return;
        }

        await ExecuteBusyActionAsync(async ct =>
        {
            var result = await WorkerService.TestTransportRuleAsync(new TestTransportRuleRequest
            {
                Sender = TestSender.Trim(),
                Recipient = TestRecipient.Trim(),
                Subject = TestSubject
            }, cancellationToken: ct);

            if (!result.IsSuccess)
            {
                SetError(result.Error?.Message ?? "Rule test error");
                return;
            }

            TestResult = result.Value == null || result.Value.MatchedRuleNames.Count == 0
                ? "No rules found"
                : string.Join(", ", result.Value.MatchedRuleNames);
        }, cancellationToken);
    }

    private void ResetEditor()
    {
        RuleIdentity = null;
        RuleName = string.Empty;
        RuleFrom = string.Empty;
        RuleSentTo = string.Empty;
        RuleSenderDomainIs = string.Empty;
        RuleRecipientDomainIs = string.Empty;
        RuleSentToMemberOf = string.Empty;
        RuleSubjectContains = string.Empty;
        RuleExceptIfFrom = string.Empty;
        RuleExceptIfSentTo = string.Empty;
        RuleExceptIfSenderDomainIs = string.Empty;
        RuleExceptIfRecipientDomainIs = string.Empty;
        RuleExceptIfSubjectContains = string.Empty;
        RulePrependSubject = string.Empty;
        RuleRedirectMessageTo = string.Empty;
        RuleBlindCopyTo = string.Empty;
        RuleAddToRecipients = string.Empty;
        RuleStopRuleProcessing = false;
        RuleDeleteMessage = false;
        RuleMode = "Enforce";
        RuleEnabled = true;
    }

    private void SetEditorProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (SetProperty(ref field, value, propertyName))
        {
            RaiseProperties(ValidationProperties);
            InvalidateCommands();
        }
    }
}
