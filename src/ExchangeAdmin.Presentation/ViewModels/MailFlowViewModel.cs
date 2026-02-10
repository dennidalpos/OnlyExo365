using System.Collections.ObjectModel;
using System.Text.RegularExpressions;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;
using ExchangeAdmin.Contracts.Messages;

namespace ExchangeAdmin.Presentation.ViewModels;

public class MailFlowViewModel : ViewModelBase
{
    private static readonly Regex EmailRegex = new(@"^[^@\s]+@[^@\s]+\.[^@\s]+$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex DomainRegex = new(@"^[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)+$", RegexOptions.Compiled | RegexOptions.CultureInvariant);

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
        SaveRuleCommand = new AsyncRelayCommand(SaveRuleAsync, () => !IsLoading && IsRuleInputValid);
        RemoveRuleCommand = new AsyncRelayCommand(RemoveRuleAsync, () => !IsLoading && SelectedRule != null);
        TestRuleCommand = new AsyncRelayCommand(TestRuleAsync, () => !IsLoading && IsRuleTestInputValid);

        SaveConnectorCommand = new AsyncRelayCommand(SaveConnectorAsync, () => !IsLoading && IsConnectorInputValid);
        RemoveConnectorCommand = new AsyncRelayCommand(RemoveConnectorAsync, () => !IsLoading && SelectedConnector != null);
        SaveDomainCommand = new AsyncRelayCommand(SaveDomainAsync, () => !IsLoading && IsDomainInputValid);
        RemoveDomainCommand = new AsyncRelayCommand(RemoveDomainAsync, () => !IsLoading && SelectedDomain != null);
    }

    public IReadOnlyList<string> RuleModes { get; } = new[] { "Enforce", "Audit" };
    public IReadOnlyList<string> ConnectorTypes { get; } = new[] { "Inbound", "Outbound" };
    public IReadOnlyList<string> DomainTypes { get; } = new[] { "Authoritative", "InternalRelay", "ExternalRelay" };

    public bool IsLoading
    {
        get => _isLoading;
        private set
        {
            if (SetProperty(ref _isLoading, value))
            {
                OnPropertyChanged(nameof(CanRefresh));
                OnPropertyChanged(nameof(CanEditSelectedRule));
                OnPropertyChanged(nameof(CanEditSelectedConnector));
                OnPropertyChanged(nameof(CanEditSelectedDomain));
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
    public bool CanEditSelectedRule => SelectedRule != null && !IsLoading;
    public bool CanEditSelectedConnector => SelectedConnector != null && !IsLoading;
    public bool CanEditSelectedDomain => SelectedDomain != null && !IsLoading;

    public bool IsRuleInputValid => !string.IsNullOrWhiteSpace(RuleName) && RuleModes.Contains(RuleMode);
    public bool IsRuleTestInputValid => IsValidEmail(TestSender) && IsValidEmail(TestRecipient);
    public bool IsConnectorInputValid => !string.IsNullOrWhiteSpace(ConnectorName) && ConnectorTypes.Contains(ConnectorType) && AreValidDomains(SplitCsv(ConnectorSenderDomains)) && AreValidDomains(SplitCsv(ConnectorRecipientDomains));
    public bool IsDomainInputValid => !string.IsNullOrWhiteSpace(DomainName) && IsValidDomain(DomainFqdn) && DomainTypes.Contains(DomainType);

    public string RuleValidationMessage => IsRuleInputValid ? string.Empty : "Rule: Name obbligatorio e Mode valido (Enforce/Audit).";
    public string TestValidationMessage => IsRuleTestInputValid ? string.Empty : "Test: Sender e Recipient devono essere email valide.";
    public string ConnectorValidationMessage => IsConnectorInputValid ? string.Empty : "Connector: Name obbligatorio, Type valido e domini in formato corretto.";
    public string DomainValidationMessage => IsDomainInputValid ? string.Empty : "Domain: Name obbligatorio, FQDN valido e DomainType supportato.";

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
                OnPropertyChanged(nameof(CanEditSelectedRule));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public ConnectorDto? SelectedConnector
    {
        get => _selectedConnector;
        set
        {
            if (SetProperty(ref _selectedConnector, value))
            {
                if (value != null)
                {
                    ConnectorIdentity = value.Identity;
                    ConnectorIdentityDisplay = string.IsNullOrWhiteSpace(value.DisplayLabel) ? value.Name : value.DisplayLabel;
                    ConnectorType = string.IsNullOrWhiteSpace(value.Type) ? "Inbound" : value.Type;
                    ConnectorName = value.Name;
                    ConnectorComment = value.Comment;
                    ConnectorEnabled = value.Enabled;
                    ConnectorSenderDomains = string.Join(",", value.SenderDomains);
                    ConnectorRecipientDomains = string.Join(",", value.RecipientDomains);
                }
                else
                {
                    ConnectorIdentity = null;
                    ConnectorIdentityDisplay = null;
                }

                OnPropertyChanged(nameof(CanEditSelectedConnector));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public AcceptedDomainDto? SelectedDomain
    {
        get => _selectedDomain;
        set
        {
            if (SetProperty(ref _selectedDomain, value))
            {
                if (value != null)
                {
                    DomainIdentity = value.Identity;
                    DomainName = value.Name;
                    DomainFqdn = value.DomainName;
                    DomainType = value.DomainType;
                    DomainMakeDefault = value.Default;
                }

                OnPropertyChanged(nameof(CanEditSelectedDomain));
                CommandManager.InvalidateRequerySuggested();
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
    private string? _connectorIdentityDisplay;
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
    public string RuleName { get => _ruleName; set { if (SetProperty(ref _ruleName, value)) RaiseValidationChanged(); } }
    public string RuleFrom { get => _ruleFrom; set => SetProperty(ref _ruleFrom, value); }
    public string RuleSentTo { get => _ruleSentTo; set => SetProperty(ref _ruleSentTo, value); }
    public string RuleSubjectContains { get => _ruleSubjectContains; set => SetProperty(ref _ruleSubjectContains, value); }
    public string RulePrependSubject { get => _rulePrependSubject; set => SetProperty(ref _rulePrependSubject, value); }
    public string RuleMode { get => _ruleMode; set { if (SetProperty(ref _ruleMode, value)) RaiseValidationChanged(); } }
    public bool RuleEnabled { get => _ruleEnabled; set => SetProperty(ref _ruleEnabled, value); }

    public string TestSender { get => _testSender; set { if (SetProperty(ref _testSender, value)) RaiseValidationChanged(); } }
    public string TestRecipient { get => _testRecipient; set { if (SetProperty(ref _testRecipient, value)) RaiseValidationChanged(); } }
    public string TestSubject { get => _testSubject; set => SetProperty(ref _testSubject, value); }
    public string TestResult { get => _testResult; set => SetProperty(ref _testResult, value); }

    public string? ConnectorIdentity { get => _connectorIdentity; set => SetProperty(ref _connectorIdentity, value); }
    public string? ConnectorIdentityDisplay { get => _connectorIdentityDisplay; set => SetProperty(ref _connectorIdentityDisplay, value); }
    public string ConnectorType { get => _connectorType; set { if (SetProperty(ref _connectorType, value)) RaiseValidationChanged(); } }
    public string ConnectorName { get => _connectorName; set { if (SetProperty(ref _connectorName, value)) RaiseValidationChanged(); } }
    public string ConnectorComment { get => _connectorComment; set => SetProperty(ref _connectorComment, value); }
    public bool ConnectorEnabled { get => _connectorEnabled; set => SetProperty(ref _connectorEnabled, value); }
    public string ConnectorSenderDomains { get => _connectorSenderDomains; set { if (SetProperty(ref _connectorSenderDomains, value)) RaiseValidationChanged(); } }
    public string ConnectorRecipientDomains { get => _connectorRecipientDomains; set { if (SetProperty(ref _connectorRecipientDomains, value)) RaiseValidationChanged(); } }

    public string? DomainIdentity { get => _domainIdentity; set => SetProperty(ref _domainIdentity, value); }
    public string DomainName { get => _domainName; set { if (SetProperty(ref _domainName, value)) RaiseValidationChanged(); } }
    public string DomainFqdn { get => _domainFqdn; set { if (SetProperty(ref _domainFqdn, value)) RaiseValidationChanged(); } }
    public string DomainType { get => _domainType; set { if (SetProperty(ref _domainType, value)) RaiseValidationChanged(); } }
    public bool DomainMakeDefault { get => _domainMakeDefault; set => SetProperty(ref _domainMakeDefault, value); }

    public ObservableCollection<TransportRuleDto> TransportRules { get; } = new();
    public ObservableCollection<ConnectorDto> Connectors { get; } = new();
    public ObservableCollection<AcceptedDomainDto> AcceptedDomains { get; } = new();

    public ICommand RefreshCommand { get; }
    public ICommand EnableRuleCommand { get; }
    public ICommand DisableRuleCommand { get; }
    public ICommand SaveRuleCommand { get; }
    public ICommand RemoveRuleCommand { get; }
    public ICommand TestRuleCommand { get; }
    public ICommand SaveConnectorCommand { get; }
    public ICommand RemoveConnectorCommand { get; }
    public ICommand SaveDomainCommand { get; }
    public ICommand RemoveDomainCommand { get; }

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
                _shellViewModel.AddLog(LogLevel.Error, $"MailFlow refresh failed: {ErrorMessage}", "MailFlow");
                return;
            }

            TransportRules.Clear();
            foreach (var item in rules.Value?.Rules ?? new List<TransportRuleDto>()) TransportRules.Add(item);
            Connectors.Clear();
            foreach (var item in connectors.Value?.Connectors ?? new List<ConnectorDto>()) Connectors.Add(item);
            AcceptedDomains.Clear();
            foreach (var item in domains.Value?.Domains ?? new List<AcceptedDomainDto>()) AcceptedDomains.Add(item);
            _shellViewModel.AddLog(LogLevel.Information, $"MailFlow refresh complete: rules={TransportRules.Count}, connectors={Connectors.Count}, domains={AcceptedDomains.Count}", "MailFlow");
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            _shellViewModel.AddLog(LogLevel.Error, $"MailFlow refresh exception: {ex.Message}", "MailFlow");
        }
        finally
        {
            IsLoading = false;
        }
    }

    private async Task SetRuleStateAsync(bool enabled)
    {
        if (SelectedRule == null) return;

        if (!enabled)
        {
            var confirmed = ErrorDialogService.ShowConfirmation("Conferma disabilitazione tenant-wide", $"Operazione: Disabilitazione regola di transport\nTarget: {SelectedRule.Name}\nImpatto: può fermare controlli/compliance tenant-wide.\n\nConfermare?");
            if (!confirmed)
            {
                return;
            }
        }

        IsLoading = true;
        ErrorMessage = null;
        try
        {
            var result = await _workerService.SetTransportRuleStateAsync(new SetTransportRuleStateRequest { Identity = SelectedRule.Identity, Enabled = enabled });
            if (!result.IsSuccess)
            {
                ErrorMessage = result.Error?.Message ?? "Errore stato rule";
                _shellViewModel.AddLog(LogLevel.Error, $"Set rule state failed (rule={SelectedRule.Name}): {ErrorMessage}", "MailFlow");
                return;
            }
            await RefreshAsync(CancellationToken.None);
        }
        finally { IsLoading = false; }
    }

    private async Task SaveRuleAsync(CancellationToken cancellationToken)
    {
        if (!IsRuleInputValid) return;

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

            if (!result.IsSuccess) { var correlationId = Guid.NewGuid().ToString("N"); ErrorMessage = $"Salvataggio regola non riuscito. Riprova o controlla i log (ref: {correlationId})."; _shellViewModel.AddLog(LogLevel.Error, $"[{correlationId}] Save rule failed (name={RuleName}): {result.Error?.Message}", "MailFlow"); return; }
            await RefreshAsync(cancellationToken);
        }
        catch (Exception ex) { ErrorMessage = ex.Message; }
        finally { IsLoading = false; }
    }

    private async Task RemoveRuleAsync(CancellationToken cancellationToken)
    {
        if (SelectedRule == null) return;
        var confirmed = ErrorDialogService.ShowConfirmation("Conferma eliminazione tenant-wide", $"Operazione: Eliminazione regola di transport\nTarget: {SelectedRule.Name}\nImpatto: rimozione permanente con impatto potenziale tenant-wide.\n\nConfermare?");
        if (!confirmed) return;

        IsLoading = true;
        ErrorMessage = null;
        try
        {
            var result = await _workerService.RemoveTransportRuleAsync(new RemoveTransportRuleRequest { Identity = SelectedRule.Identity }, cancellationToken: cancellationToken);
            if (!result.IsSuccess) { ErrorMessage = result.Error?.Message ?? "Errore eliminazione rule"; _shellViewModel.AddLog(LogLevel.Error, $"Remove rule failed (rule={SelectedRule.Name}): {ErrorMessage}", "MailFlow"); return; }
            await RefreshAsync(cancellationToken);
        }
        catch (Exception ex) { ErrorMessage = ex.Message; }
        finally { IsLoading = false; }
    }

    private async Task TestRuleAsync(CancellationToken cancellationToken)
    {
        if (!IsRuleTestInputValid) return;

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
        if (!IsConnectorInputValid) return;

        if (SelectedConnector != null && SelectedConnector.Enabled && !ConnectorEnabled)
        {
            var disableConfirmed = ErrorDialogService.ShowConfirmation(
                "Conferma disabilitazione tenant-wide",
                $"Operazione: Disabilitazione connector {SelectedConnector.Type}
Target: {SelectedConnector.Name}
Impatto: può interrompere il flusso posta tenant-wide.

Confermare?");
            if (!disableConfirmed)
            {
                return;
            }
        }

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

            if (!result.IsSuccess) { var correlationId = Guid.NewGuid().ToString("N"); ErrorMessage = $"Salvataggio connector non riuscito. Verificare i campi e riprovare (ref: {correlationId})."; _shellViewModel.AddLog(LogLevel.Error, $"[{correlationId}] Save connector failed (name={ConnectorName}, type={ConnectorType}): {result.Error?.Message}", "MailFlow"); return; }
            await RefreshAsync(cancellationToken);
        }
        catch (Exception ex) { var correlationId = Guid.NewGuid().ToString("N"); ErrorMessage = $"Errore durante il salvataggio connector (ref: {correlationId})."; _shellViewModel.AddLog(LogLevel.Error, $"[{correlationId}] Save connector exception: {ex.Message}", "MailFlow"); }
        finally { IsLoading = false; }
    }

    private async Task RemoveConnectorAsync(CancellationToken cancellationToken)
    {
        if (SelectedConnector == null) return;
        var confirmed = ErrorDialogService.ShowConfirmation("Conferma eliminazione tenant-wide", $"Operazione: Eliminazione connector {SelectedConnector.Type}\nTarget: {SelectedConnector.Name}\nImpatto: modifica routing posta tenant-wide.\n\nConfermare?");
        if (!confirmed) return;

        IsLoading = true;
        ErrorMessage = null;
        try
        {
            var result = await _workerService.RemoveConnectorAsync(new RemoveConnectorRequest
            {
                Identity = SelectedConnector.Identity,
                Type = SelectedConnector.Type
            }, cancellationToken: cancellationToken);

            if (!result.IsSuccess) { ErrorMessage = result.Error?.Message ?? "Errore eliminazione connector"; _shellViewModel.AddLog(LogLevel.Error, $"Remove connector failed (connector={SelectedConnector.Name}): {ErrorMessage}", "MailFlow"); return; }
            await RefreshAsync(cancellationToken);
        }
        catch (Exception ex) { ErrorMessage = ex.Message; }
        finally { IsLoading = false; }
    }

    private async Task SaveDomainAsync(CancellationToken cancellationToken)
    {
        if (!IsDomainInputValid) return;

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

            if (!result.IsSuccess) { var correlationId = Guid.NewGuid().ToString("N"); ErrorMessage = $"Salvataggio dominio non riuscito (ref: {correlationId})."; _shellViewModel.AddLog(LogLevel.Error, $"[{correlationId}] Save domain failed (domain={DomainFqdn}): {result.Error?.Message}", "MailFlow"); return; }
            await RefreshAsync(cancellationToken);
        }
        catch (Exception ex) { var correlationId = Guid.NewGuid().ToString("N"); ErrorMessage = $"Errore durante il salvataggio dominio (ref: {correlationId})."; _shellViewModel.AddLog(LogLevel.Error, $"[{correlationId}] Save domain exception: {ex.Message}", "MailFlow"); }
        finally { IsLoading = false; }
    }

    private async Task RemoveDomainAsync(CancellationToken cancellationToken)
    {
        if (SelectedDomain == null) return;
        var confirmed = ErrorDialogService.ShowConfirmation("Conferma eliminazione tenant-wide", $"Operazione: Eliminazione accepted domain\nTarget: {SelectedDomain.DomainName}\nImpatto: può interrompere recapito/routing tenant-wide.\n\nConfermare?");
        if (!confirmed) return;

        IsLoading = true;
        ErrorMessage = null;
        try
        {
            var result = await _workerService.RemoveAcceptedDomainAsync(new RemoveAcceptedDomainRequest { Identity = SelectedDomain.Identity }, cancellationToken: cancellationToken);
            if (!result.IsSuccess) { ErrorMessage = result.Error?.Message ?? "Errore eliminazione dominio"; _shellViewModel.AddLog(LogLevel.Error, $"Remove domain failed (domain={SelectedDomain.DomainName}): {ErrorMessage}", "MailFlow"); return; }
            await RefreshAsync(cancellationToken);
        }
        catch (Exception ex) { ErrorMessage = ex.Message; }
        finally { IsLoading = false; }
    }

    private void RaiseValidationChanged()
    {
        OnPropertyChanged(nameof(IsRuleInputValid));
        OnPropertyChanged(nameof(IsRuleTestInputValid));
        OnPropertyChanged(nameof(IsConnectorInputValid));
        OnPropertyChanged(nameof(IsDomainInputValid));
        OnPropertyChanged(nameof(RuleValidationMessage));
        OnPropertyChanged(nameof(TestValidationMessage));
        OnPropertyChanged(nameof(ConnectorValidationMessage));
        OnPropertyChanged(nameof(DomainValidationMessage));
        CommandManager.InvalidateRequerySuggested();
    }

    private static List<string> SplitCsv(string value) => value
        .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
        .Where(v => !string.IsNullOrWhiteSpace(v))
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToList();

    private static bool IsValidEmail(string value) => !string.IsNullOrWhiteSpace(value) && EmailRegex.IsMatch(value.Trim());

    private static bool IsValidDomain(string value) => !string.IsNullOrWhiteSpace(value) && DomainRegex.IsMatch(value.Trim());

    private static bool AreValidDomains(IEnumerable<string> domains) => domains.All(IsValidDomain);
}
