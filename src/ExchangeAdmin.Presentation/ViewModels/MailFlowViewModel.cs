using System.Collections.ObjectModel;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

public class MailFlowViewModel : ViewModelBase
{
    private readonly IWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;

    private bool _isLoading;
    private string? _errorMessage;
    private TransportRuleDto? _selectedRule;
    private ConnectorDto? _selectedConnector;
    private AcceptedDomainDto? _selectedDomain;

    public MailFlowViewModel(IWorkerService workerService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;

        RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => CanRefresh);

        EnableRuleCommand = new AsyncRelayCommand(() => SetRuleStateAsync(true), () => SelectedRule != null && !IsLoading);
        DisableRuleCommand = new AsyncRelayCommand(() => SetRuleStateAsync(false), () => SelectedRule != null && !IsLoading);
        SaveRuleCommand = new AsyncRelayCommand(SaveRuleAsync, () => !IsLoading && !string.IsNullOrWhiteSpace(RuleName));
        TestRuleCommand = new AsyncRelayCommand(TestRuleAsync, () => !IsLoading && !string.IsNullOrWhiteSpace(TestSender) && !string.IsNullOrWhiteSpace(TestRecipient));

        SaveConnectorCommand = new AsyncRelayCommand(SaveConnectorAsync, () => !IsLoading && !string.IsNullOrWhiteSpace(ConnectorName));
        SaveDomainCommand = new AsyncRelayCommand(SaveDomainAsync, () => !IsLoading && !string.IsNullOrWhiteSpace(DomainName) && !string.IsNullOrWhiteSpace(DomainFqdn));
    }

    public bool IsLoading
    {
        get => _isLoading;
        private set
        {
            if (SetProperty(ref _isLoading, value))
            {
                OnPropertyChanged(nameof(CanRefresh));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public string? ErrorMessage
    {
        get => _errorMessage;
        private set
        {
            if (SetProperty(ref _errorMessage, value))
            {
                OnPropertyChanged(nameof(HasError));
            }
        }
    }

    public bool HasError => !string.IsNullOrWhiteSpace(ErrorMessage);
    public bool CanRefresh => !IsLoading && _shellViewModel.IsExchangeConnected;

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
                    RuleSubjectContains = string.Join(",", value.SubjectContainsWords);
                    RulePrependSubject = value.PrependSubject;
                    RuleMode = string.IsNullOrWhiteSpace(value.Mode) ? "Enforce" : value.Mode;
                    RuleEnabled = !string.Equals(value.State, "Disabled", StringComparison.OrdinalIgnoreCase);
                }
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public ConnectorDto? SelectedConnector
    {
        get => _selectedConnector;
        set
        {
            if (SetProperty(ref _selectedConnector, value) && value != null)
            {
                ConnectorIdentity = value.Identity;
                ConnectorType = string.IsNullOrWhiteSpace(value.Type) ? "Inbound" : value.Type;
                ConnectorName = value.Name;
                ConnectorComment = value.Comment;
                ConnectorEnabled = value.Enabled;
                ConnectorSenderDomains = string.Join(",", value.SenderDomains);
                ConnectorRecipientDomains = string.Join(",", value.RecipientDomains);
            }
        }
    }

    public AcceptedDomainDto? SelectedDomain
    {
        get => _selectedDomain;
        set
        {
            if (SetProperty(ref _selectedDomain, value) && value != null)
            {
                DomainIdentity = value.Identity;
                DomainName = value.Name;
                DomainFqdn = value.DomainName;
                DomainType = value.DomainType;
                DomainMakeDefault = value.Default;
            }
        }
    }

    private string? _ruleIdentity;
    private string _ruleName = string.Empty;
    private string _ruleFrom = string.Empty;
    private string _ruleSentTo = string.Empty;
    private string _ruleSubjectContains = string.Empty;
    private string _rulePrependSubject = string.Empty;
    private string _ruleMode = "Enforce";
    private bool _ruleEnabled = true;
    private string _testSender = string.Empty;
    private string _testRecipient = string.Empty;
    private string _testSubject = string.Empty;
    private string _testResult = string.Empty;

    private string? _connectorIdentity;
    private string _connectorType = "Inbound";
    private string _connectorName = string.Empty;
    private string _connectorComment = string.Empty;
    private bool _connectorEnabled = true;
    private string _connectorSenderDomains = string.Empty;
    private string _connectorRecipientDomains = string.Empty;

    private string? _domainIdentity;
    private string _domainName = string.Empty;
    private string _domainFqdn = string.Empty;
    private string _domainType = "Authoritative";
    private bool _domainMakeDefault;

    public string? RuleIdentity { get => _ruleIdentity; set => SetProperty(ref _ruleIdentity, value); }
    public string RuleName { get => _ruleName; set => SetProperty(ref _ruleName, value); }
    public string RuleFrom { get => _ruleFrom; set => SetProperty(ref _ruleFrom, value); }
    public string RuleSentTo { get => _ruleSentTo; set => SetProperty(ref _ruleSentTo, value); }
    public string RuleSubjectContains { get => _ruleSubjectContains; set => SetProperty(ref _ruleSubjectContains, value); }
    public string RulePrependSubject { get => _rulePrependSubject; set => SetProperty(ref _rulePrependSubject, value); }
    public string RuleMode { get => _ruleMode; set => SetProperty(ref _ruleMode, value); }
    public bool RuleEnabled { get => _ruleEnabled; set => SetProperty(ref _ruleEnabled, value); }
    public string TestSender { get => _testSender; set => SetProperty(ref _testSender, value); }
    public string TestRecipient { get => _testRecipient; set => SetProperty(ref _testRecipient, value); }
    public string TestSubject { get => _testSubject; set => SetProperty(ref _testSubject, value); }
    public string TestResult { get => _testResult; set => SetProperty(ref _testResult, value); }

    public string? ConnectorIdentity { get => _connectorIdentity; set => SetProperty(ref _connectorIdentity, value); }
    public string ConnectorType { get => _connectorType; set => SetProperty(ref _connectorType, value); }
    public string ConnectorName { get => _connectorName; set => SetProperty(ref _connectorName, value); }
    public string ConnectorComment { get => _connectorComment; set => SetProperty(ref _connectorComment, value); }
    public bool ConnectorEnabled { get => _connectorEnabled; set => SetProperty(ref _connectorEnabled, value); }
    public string ConnectorSenderDomains { get => _connectorSenderDomains; set => SetProperty(ref _connectorSenderDomains, value); }
    public string ConnectorRecipientDomains { get => _connectorRecipientDomains; set => SetProperty(ref _connectorRecipientDomains, value); }

    public string? DomainIdentity { get => _domainIdentity; set => SetProperty(ref _domainIdentity, value); }
    public string DomainName { get => _domainName; set => SetProperty(ref _domainName, value); }
    public string DomainFqdn { get => _domainFqdn; set => SetProperty(ref _domainFqdn, value); }
    public string DomainType { get => _domainType; set => SetProperty(ref _domainType, value); }
    public bool DomainMakeDefault { get => _domainMakeDefault; set => SetProperty(ref _domainMakeDefault, value); }

    public ObservableCollection<TransportRuleDto> TransportRules { get; } = new();
    public ObservableCollection<ConnectorDto> Connectors { get; } = new();
    public ObservableCollection<AcceptedDomainDto> AcceptedDomains { get; } = new();

    public ICommand RefreshCommand { get; }
    public ICommand EnableRuleCommand { get; }
    public ICommand DisableRuleCommand { get; }
    public ICommand SaveRuleCommand { get; }
    public ICommand TestRuleCommand { get; }
    public ICommand SaveConnectorCommand { get; }
    public ICommand SaveDomainCommand { get; }

    public async Task LoadAsync()
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            ErrorMessage = "Non connesso a Exchange Online";
            return;
        }

        if (TransportRules.Count == 0 && Connectors.Count == 0 && AcceptedDomains.Count == 0)
        {
            await RefreshAsync(CancellationToken.None);
        }
    }

    private async Task RefreshAsync(CancellationToken cancellationToken)
    {
        IsLoading = true;
        ErrorMessage = null;

        try
        {
            var rulesTask = _workerService.GetTransportRulesAsync(new GetTransportRulesRequest(), cancellationToken: cancellationToken);
            var connectorsTask = _workerService.GetConnectorsAsync(new GetConnectorsRequest(), cancellationToken: cancellationToken);
            var domainsTask = _workerService.GetAcceptedDomainsAsync(new GetAcceptedDomainsRequest(), cancellationToken: cancellationToken);
            await Task.WhenAll(rulesTask, connectorsTask, domainsTask);

            var rules = await rulesTask;
            var connectors = await connectorsTask;
            var domains = await domainsTask;

            if (!rules.IsSuccess || !connectors.IsSuccess || !domains.IsSuccess)
            {
                ErrorMessage = rules.Error?.Message ?? connectors.Error?.Message ?? domains.Error?.Message ?? "Errore caricamento mail flow";
                return;
            }

            TransportRules.Clear();
            foreach (var item in rules.Value?.Rules ?? new List<TransportRuleDto>()) TransportRules.Add(item);
            Connectors.Clear();
            foreach (var item in connectors.Value?.Connectors ?? new List<ConnectorDto>()) Connectors.Add(item);
            AcceptedDomains.Clear();
            foreach (var item in domains.Value?.Domains ?? new List<AcceptedDomainDto>()) AcceptedDomains.Add(item);
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
        finally
        {
            IsLoading = false;
        }
    }

    private async Task SetRuleStateAsync(bool enabled)
    {
        if (SelectedRule == null) return;
        IsLoading = true;
        ErrorMessage = null;
        try
        {
            var result = await _workerService.SetTransportRuleStateAsync(new SetTransportRuleStateRequest { Identity = SelectedRule.Identity, Enabled = enabled });
            if (!result.IsSuccess) { ErrorMessage = result.Error?.Message ?? "Errore stato rule"; return; }
            await RefreshAsync(CancellationToken.None);
        }
        finally { IsLoading = false; }
    }

    private async Task SaveRuleAsync(CancellationToken cancellationToken)
    {
        IsLoading = true;
        ErrorMessage = null;
        try
        {
            var result = await _workerService.UpsertTransportRuleAsync(new UpsertTransportRuleRequest
            {
                Identity = string.IsNullOrWhiteSpace(RuleIdentity) ? null : RuleIdentity,
                Name = RuleName.Trim(),
                From = SplitCsv(RuleFrom),
                SentTo = SplitCsv(RuleSentTo),
                SubjectContainsWords = SplitCsv(RuleSubjectContains),
                PrependSubject = string.IsNullOrWhiteSpace(RulePrependSubject) ? null : RulePrependSubject.Trim(),
                Mode = RuleMode,
                Enabled = RuleEnabled
            }, cancellationToken: cancellationToken);

            if (!result.IsSuccess) { ErrorMessage = result.Error?.Message ?? "Errore salvataggio rule"; return; }
            await RefreshAsync(cancellationToken);
        }
        catch (Exception ex) { ErrorMessage = ex.Message; }
        finally { IsLoading = false; }
    }

    private async Task TestRuleAsync(CancellationToken cancellationToken)
    {
        IsLoading = true;
        ErrorMessage = null;
        try
        {
            var result = await _workerService.TestTransportRuleAsync(new TestTransportRuleRequest
            {
                Sender = TestSender.Trim(),
                Recipient = TestRecipient.Trim(),
                Subject = TestSubject
            }, cancellationToken: cancellationToken);

            if (!result.IsSuccess) { ErrorMessage = result.Error?.Message ?? "Errore test rule"; return; }
            TestResult = result.Value == null || result.Value.MatchedRuleNames.Count == 0
                ? "Nessuna regola trovata"
                : string.Join(", ", result.Value.MatchedRuleNames);
        }
        catch (Exception ex) { ErrorMessage = ex.Message; }
        finally { IsLoading = false; }
    }

    private async Task SaveConnectorAsync(CancellationToken cancellationToken)
    {
        IsLoading = true;
        ErrorMessage = null;
        try
        {
            var result = await _workerService.UpsertConnectorAsync(new UpsertConnectorRequest
            {
                Identity = string.IsNullOrWhiteSpace(ConnectorIdentity) ? null : ConnectorIdentity,
                Type = ConnectorType,
                Name = ConnectorName.Trim(),
                Comment = ConnectorComment,
                Enabled = ConnectorEnabled,
                SenderDomains = SplitCsv(ConnectorSenderDomains),
                RecipientDomains = SplitCsv(ConnectorRecipientDomains)
            }, cancellationToken: cancellationToken);

            if (!result.IsSuccess) { ErrorMessage = result.Error?.Message ?? "Errore salvataggio connector"; return; }
            await RefreshAsync(cancellationToken);
        }
        catch (Exception ex) { ErrorMessage = ex.Message; }
        finally { IsLoading = false; }
    }

    private async Task SaveDomainAsync(CancellationToken cancellationToken)
    {
        IsLoading = true;
        ErrorMessage = null;
        try
        {
            var result = await _workerService.UpsertAcceptedDomainAsync(new UpsertAcceptedDomainRequest
            {
                Identity = string.IsNullOrWhiteSpace(DomainIdentity) ? null : DomainIdentity,
                Name = DomainName.Trim(),
                DomainName = DomainFqdn.Trim(),
                DomainType = DomainType,
                MakeDefault = DomainMakeDefault
            }, cancellationToken: cancellationToken);

            if (!result.IsSuccess) { ErrorMessage = result.Error?.Message ?? "Errore salvataggio dominio"; return; }
            await RefreshAsync(cancellationToken);
        }
        catch (Exception ex) { ErrorMessage = ex.Message; }
        finally { IsLoading = false; }
    }

    private static List<string> SplitCsv(string value) => value
        .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
        .Where(v => !string.IsNullOrWhiteSpace(v))
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToList();
}
