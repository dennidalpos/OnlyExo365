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

    public MailFlowViewModel(IWorkerService workerService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;

        RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => CanRefresh);
        EnableRuleCommand = new AsyncRelayCommand(() => SetRuleStateAsync(true), () => SelectedRule != null && !IsLoading);
        DisableRuleCommand = new AsyncRelayCommand(() => SetRuleStateAsync(false), () => SelectedRule != null && !IsLoading);
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
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public ObservableCollection<TransportRuleDto> TransportRules { get; } = new();
    public ObservableCollection<ConnectorDto> Connectors { get; } = new();
    public ObservableCollection<AcceptedDomainDto> AcceptedDomains { get; } = new();

    public ICommand RefreshCommand { get; }
    public ICommand EnableRuleCommand { get; }
    public ICommand DisableRuleCommand { get; }

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

            if (!rules.IsSuccess)
            {
                ErrorMessage = rules.Error?.Message ?? "Errore caricamento transport rules";
                return;
            }

            if (!connectors.IsSuccess)
            {
                ErrorMessage = connectors.Error?.Message ?? "Errore caricamento connectors";
                return;
            }

            if (!domains.IsSuccess)
            {
                ErrorMessage = domains.Error?.Message ?? "Errore caricamento accepted domains";
                return;
            }

            TransportRules.Clear();
            foreach (var rule in rules.Value?.Rules ?? new List<TransportRuleDto>())
            {
                TransportRules.Add(rule);
            }

            Connectors.Clear();
            foreach (var connector in connectors.Value?.Connectors ?? new List<ConnectorDto>())
            {
                Connectors.Add(connector);
            }

            AcceptedDomains.Clear();
            foreach (var domain in domains.Value?.Domains ?? new List<AcceptedDomainDto>())
            {
                AcceptedDomains.Add(domain);
            }

            _shellViewModel.AddLog(LogLevel.Information, $"Mail flow loaded: {TransportRules.Count} rules, {Connectors.Count} connectors, {AcceptedDomains.Count} domains");
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
        if (SelectedRule == null)
        {
            return;
        }

        IsLoading = true;
        ErrorMessage = null;

        try
        {
            var result = await _workerService.SetTransportRuleStateAsync(new SetTransportRuleStateRequest
            {
                Identity = SelectedRule.Identity,
                Enabled = enabled
            });

            if (!result.IsSuccess)
            {
                ErrorMessage = result.Error?.Message ?? "Errore aggiornamento stato rule";
                return;
            }

            await RefreshAsync(CancellationToken.None);
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
}
