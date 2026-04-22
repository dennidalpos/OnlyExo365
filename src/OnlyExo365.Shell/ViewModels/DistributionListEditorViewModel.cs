using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Shell.Helpers;

namespace OnlyExo365.Shell.ViewModels;

public sealed class DistributionListEditorViewModel : ViewModelBase
{
    private readonly IDistributionListsWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;
    private readonly Action<string?> _setErrorMessage;
    private readonly Func<CancellationToken, Task> _refreshAsync;

    private string? _newDistributionListDisplayName;
    private string? _newDistributionListAlias;
    private string? _newDistributionListLocalPart;
    private string? _selectedDistributionListDomain;
    private bool _isCreatingDistributionList;

    public DistributionListEditorViewModel(
        IDistributionListsWorkerService workerService,
        ShellViewModel shellViewModel,
        Action<string?> setErrorMessage,
        Func<CancellationToken, Task> refreshAsync)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;
        _setErrorMessage = setErrorMessage;
        _refreshAsync = refreshAsync;

        CreateDistributionListCommand = new AsyncRelayCommand(CreateDistributionListAsync, () => CanCreateDistributionList);
    }

    public ObservableCollection<string> AvailableMailDomains { get; } = new();

    public string? NewDistributionListDisplayName
    {
        get => _newDistributionListDisplayName;
        set
        {
            if (SetProperty(ref _newDistributionListDisplayName, value))
            {
                OnPropertyChanged(nameof(CanCreateDistributionList));
            }
        }
    }

    public string? NewDistributionListAlias
    {
        get => _newDistributionListAlias;
        set
        {
            if (SetProperty(ref _newDistributionListAlias, value))
            {
                OnPropertyChanged(nameof(CanCreateDistributionList));
            }
        }
    }

    public string? NewDistributionListLocalPart
    {
        get => _newDistributionListLocalPart;
        set
        {
            if (SetProperty(ref _newDistributionListLocalPart, value))
            {
                OnPropertyChanged(nameof(CanCreateDistributionList));
            }
        }
    }

    public string? SelectedDistributionListDomain
    {
        get => _selectedDistributionListDomain;
        set
        {
            if (SetProperty(ref _selectedDistributionListDomain, value))
            {
                OnPropertyChanged(nameof(CanCreateDistributionList));
            }
        }
    }

    public bool IsCreatingDistributionList
    {
        get => _isCreatingDistributionList;
        private set
        {
            if (SetProperty(ref _isCreatingDistributionList, value))
            {
                OnPropertyChanged(nameof(CanCreateDistributionList));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool CanCreateDistributionList =>
        !IsCreatingDistributionList && _shellViewModel.IsExchangeConnected &&
        !string.IsNullOrWhiteSpace(NewDistributionListDisplayName) &&
        !string.IsNullOrWhiteSpace(NewDistributionListAlias) &&
        !string.IsNullOrWhiteSpace(NewDistributionListLocalPart) &&
        !string.IsNullOrWhiteSpace(SelectedDistributionListDomain);

    public ICommand CreateDistributionListCommand { get; }

    public async Task LoadAcceptedDomainsAsync(CancellationToken cancellationToken)
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            AvailableMailDomains.Clear();
            if (!string.IsNullOrWhiteSpace(SelectedDistributionListDomain))
            {
                SelectedDistributionListDomain = null;
            }

            return;
        }

        try
        {
            var result = await _workerService.GetAcceptedDomainsAsync(new GetAcceptedDomainsRequest(), cancellationToken: cancellationToken);
            if (!result.IsSuccess || result.Value == null)
            {
                return;
            }

            var domains = result.Value.Domains
                .Select(domain => domain.DomainName?.Trim())
                .Where(domain => !string.IsNullOrWhiteSpace(domain))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(domain => domain, StringComparer.OrdinalIgnoreCase)
                .ToList();

            AvailableMailDomains.Clear();
            foreach (var domain in domains)
            {
                AvailableMailDomains.Add(domain!);
            }

            if (string.IsNullOrWhiteSpace(SelectedDistributionListDomain))
            {
                SelectedDistributionListDomain = result.Value.Domains.FirstOrDefault(d => d.Default)?.DomainName ?? AvailableMailDomains.FirstOrDefault();
            }
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Warning, $"Unable to load accepted domains: {ex.Message}");
        }
    }

    public void HandleShellPropertyChanged()
    {
        OnPropertyChanged(nameof(CanCreateDistributionList));
        CommandManager.InvalidateRequerySuggested();
    }

    private async Task CreateDistributionListAsync(CancellationToken cancellationToken)
    {
        if (!CanCreateDistributionList)
        {
            return;
        }

        var primarySmtpAddress = $"{NewDistributionListLocalPart!.Trim()}@{SelectedDistributionListDomain!.Trim()}";
        if (!ConfirmMutation(
                "Create distribution list",
                primarySmtpAddress,
                "Create a new distribution list in the tenant.",
                "Confirm distribution list creation"))
        {
            return;
        }

        IsCreatingDistributionList = true;
        _setErrorMessage(null);

        try
        {
            var result = await _workerService.CreateDistributionListAsync(
                new CreateDistributionListRequest
                {
                    DisplayName = NewDistributionListDisplayName!.Trim(),
                    Alias = NewDistributionListAlias!.Trim(),
                    PrimarySmtpAddress = primarySmtpAddress
                },
                cancellationToken: cancellationToken);

            if (result.IsSuccess)
            {
                NewDistributionListDisplayName = string.Empty;
                NewDistributionListAlias = string.Empty;
                NewDistributionListLocalPart = string.Empty;
                await _refreshAsync(cancellationToken);
                return;
            }

            if (!result.WasCancelled)
            {
                _setErrorMessage(result.Error?.Message ?? "Unable to create the distribution list.");
            }
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            _setErrorMessage(ex.Message);
            _shellViewModel.AddLog(LogLevel.Error, $"Create distribution group error: {ex.Message}");
        }
        finally
        {
            IsCreatingDistributionList = false;
        }
    }
}

