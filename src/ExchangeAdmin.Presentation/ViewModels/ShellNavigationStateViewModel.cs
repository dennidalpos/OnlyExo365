using System.ComponentModel;
using System.Windows.Input;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;
using ExchangeAdmin.Presentation.Text;

namespace ExchangeAdmin.Presentation.ViewModels;

public sealed class ShellNavigationStateViewModel : ViewModelBase, IDisposable
{
    private readonly NavigationService _navigationService;
    private readonly ShellProgressViewModel _progressState;
    private readonly ShellPromptViewModel _promptState;
    private readonly List<NavigationBlockingRegistration> _blockingSources = new();
    private readonly List<Func<bool>> _unsavedChangesChecks = new();

    private NavigationPage _currentPage = NavigationPage.Dashboard;
    private bool _isNavigationLocked;

    public ShellNavigationStateViewModel(
        NavigationService navigationService,
        ShellProgressViewModel progressState,
        ShellPromptViewModel promptState)
    {
        _navigationService = navigationService;
        _progressState = progressState;
        _promptState = promptState;

        _navigationService.PageChanged += OnPageChanged;
        _navigationService.Navigating += OnNavigating;
        _navigationService.NavigationStateChanged += OnNavigationStateChanged;
        _progressState.PropertyChanged += OnProgressStatePropertyChanged;

        NavigateToDashboardCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.Dashboard));
        NavigateToContactsCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.Contacts));
        NavigateToResourcesCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.Resources));
        NavigateToPublicFoldersCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.PublicFolders));
        NavigateToMobileDevicesCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.MobileDevices));
        NavigateToMigrationCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.Migration));
        NavigateToPermissionsCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.Permissions));
        NavigateToMailboxesCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.Mailboxes));
        NavigateToDeletedMailboxesCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.DeletedMailboxes));
        NavigateToMailboxSpaceCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.MailboxSpace));
        NavigateToMailboxAccessReportCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.MailboxAccessReport));
        NavigateToDistributionListsCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.DistributionLists));
        NavigateToMessageTraceCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.MessageTrace));
        NavigateToComplianceCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.Compliance));
        NavigateToMailSecurityCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.MailSecurity));
        NavigateToMailFlowCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.MailFlow));
        NavigateToToolsCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.Tools));
        NavigateToLogsCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.Logs));
    }

    public NavigationPage CurrentPage
    {
        get => _currentPage;
        internal set
        {
            if (SetProperty(ref _currentPage, value))
            {
                NotifyNavigationPropertiesChanged();
            }
        }
    }

    public bool IsDashboardPage => CurrentPage == NavigationPage.Dashboard;
    public bool IsContactsPage => CurrentPage == NavigationPage.Contacts;
    public bool IsResourcesPage => CurrentPage == NavigationPage.Resources;
    public bool IsPublicFoldersPage => CurrentPage == NavigationPage.PublicFolders;
    public bool IsMobileDevicesPage => CurrentPage == NavigationPage.MobileDevices;
    public bool IsMigrationPage => CurrentPage == NavigationPage.Migration;
    public bool IsPermissionsPage => CurrentPage == NavigationPage.Permissions;
    public bool IsMailboxesPage => CurrentPage == NavigationPage.Mailboxes;
    public bool IsDeletedMailboxesPage => CurrentPage == NavigationPage.DeletedMailboxes;
    public bool IsMailboxSpacePage => CurrentPage == NavigationPage.MailboxSpace;
    public bool IsMailboxAccessReportPage => CurrentPage == NavigationPage.MailboxAccessReport;
    public bool IsDistributionListsPage => CurrentPage == NavigationPage.DistributionLists;
    public bool IsMessageTracePage => CurrentPage == NavigationPage.MessageTrace;
    public bool IsCompliancePage => CurrentPage == NavigationPage.Compliance;
    public bool IsMailSecurityPage => CurrentPage == NavigationPage.MailSecurity;
    public bool IsMailFlowPage => CurrentPage == NavigationPage.MailFlow;
    public bool IsToolsPage => CurrentPage == NavigationPage.Tools;
    public bool IsLogsPage => CurrentPage == NavigationPage.Logs;

    public string CurrentPageTitle => CurrentPage switch
    {
        _ => UiTextCatalog.GetNavigationLabel(CurrentPage)
    };

    public bool IsNavigationLocked
    {
        get => _isNavigationLocked;
        private set
        {
            if (SetProperty(ref _isNavigationLocked, value))
            {
                OnPropertyChanged(nameof(CanNavigate));
            }
        }
    }

    public bool CanNavigate => !IsNavigationLocked;

    public ICommand NavigateToDashboardCommand { get; }
    public ICommand NavigateToContactsCommand { get; }
    public ICommand NavigateToResourcesCommand { get; }
    public ICommand NavigateToPublicFoldersCommand { get; }
    public ICommand NavigateToMobileDevicesCommand { get; }
    public ICommand NavigateToMigrationCommand { get; }
    public ICommand NavigateToPermissionsCommand { get; }
    public ICommand NavigateToMailboxesCommand { get; }
    public ICommand NavigateToDeletedMailboxesCommand { get; }
    public ICommand NavigateToMailboxSpaceCommand { get; }
    public ICommand NavigateToMailboxAccessReportCommand { get; }
    public ICommand NavigateToDistributionListsCommand { get; }
    public ICommand NavigateToMessageTraceCommand { get; }
    public ICommand NavigateToComplianceCommand { get; }
    public ICommand NavigateToMailSecurityCommand { get; }
    public ICommand NavigateToMailFlowCommand { get; }
    public ICommand NavigateToToolsCommand { get; }
    public ICommand NavigateToLogsCommand { get; }

    public void RegisterBlockingStateSource(INotifyPropertyChanged source, Func<bool> isBlocking, params string[] watchedProperties)
    {
        if (_blockingSources.Any(registration => ReferenceEquals(registration.Source, source)))
        {
            return;
        }

        _blockingSources.Add(new NavigationBlockingRegistration(source, isBlocking, watchedProperties));
        source.PropertyChanged += OnBlockingSourcePropertyChanged;
        UpdateNavigationLock();
    }

    public void RegisterUnsavedChangesCheck(Func<bool> hasUnsavedChanges)
    {
        _unsavedChangesChecks.Add(hasUnsavedChanges);
    }

    public void Dispose()
    {
        _navigationService.PageChanged -= OnPageChanged;
        _navigationService.Navigating -= OnNavigating;
        _navigationService.NavigationStateChanged -= OnNavigationStateChanged;
        _progressState.PropertyChanged -= OnProgressStatePropertyChanged;

        foreach (var registration in _blockingSources)
        {
            registration.Source.PropertyChanged -= OnBlockingSourcePropertyChanged;
        }
    }

    private void OnPageChanged(object? sender, NavigationPage page)
    {
        RunOnUiThread(() => CurrentPage = page);
    }

    private void OnNavigating(object? sender, NavigatingEventArgs e)
    {
        if (IsNavigationLocked)
        {
            _promptState.ShowInformation(
                UserMessageCatalog.OperationInProgressTitle,
                UserMessageCatalog.OperationInProgressMessage);
            e.Cancel = true;
            return;
        }

        if (!HasUnsavedChanges())
        {
            return;
        }

        var confirmed = _promptState.ShowConfirmationBlocking(
            UserMessageCatalog.UnsavedChangesTitle,
            UserMessageCatalog.UnsavedChangesMessage);

        if (!confirmed)
        {
            e.Cancel = true;
        }
    }

    private void OnNavigationStateChanged(object? sender, EventArgs e)
    {
        UpdateNavigationLock();
    }

    private bool HasUnsavedChanges()
    {
        foreach (var check in _unsavedChangesChecks)
        {
            if (check())
            {
                return true;
            }
        }

        return false;
    }

    private void OnProgressStatePropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(e.PropertyName) ||
            e.PropertyName == nameof(ShellProgressViewModel.IsGlobalOperationRunning))
        {
            UpdateNavigationLock();
        }
    }

    private void OnBlockingSourcePropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        var registration = _blockingSources.FirstOrDefault(item => ReferenceEquals(item.Source, sender));
        if (registration == null || !registration.ShouldRecompute(e.PropertyName))
        {
            return;
        }

        UpdateNavigationLock();
    }

    private void UpdateNavigationLock()
    {
        IsNavigationLocked =
            _navigationService.IsNavigationPending ||
            _progressState.IsGlobalOperationRunning ||
            _blockingSources.Any(registration => registration.IsBlocking());
    }

    private void NotifyNavigationPropertiesChanged()
    {
        OnPropertyChanged(nameof(IsDashboardPage));
        OnPropertyChanged(nameof(IsContactsPage));
        OnPropertyChanged(nameof(IsResourcesPage));
        OnPropertyChanged(nameof(IsPublicFoldersPage));
        OnPropertyChanged(nameof(IsMobileDevicesPage));
        OnPropertyChanged(nameof(IsMigrationPage));
        OnPropertyChanged(nameof(IsPermissionsPage));
        OnPropertyChanged(nameof(IsMailboxesPage));
        OnPropertyChanged(nameof(IsDeletedMailboxesPage));
        OnPropertyChanged(nameof(IsMailboxSpacePage));
        OnPropertyChanged(nameof(IsMailboxAccessReportPage));
        OnPropertyChanged(nameof(IsDistributionListsPage));
        OnPropertyChanged(nameof(IsMessageTracePage));
        OnPropertyChanged(nameof(IsCompliancePage));
        OnPropertyChanged(nameof(IsMailSecurityPage));
        OnPropertyChanged(nameof(IsMailFlowPage));
        OnPropertyChanged(nameof(IsToolsPage));
        OnPropertyChanged(nameof(IsLogsPage));
        OnPropertyChanged(nameof(CurrentPageTitle));
    }

    private sealed class NavigationBlockingRegistration
    {
        private readonly HashSet<string>? _watchedProperties;

        public NavigationBlockingRegistration(INotifyPropertyChanged source, Func<bool> isBlocking, IEnumerable<string> watchedProperties)
        {
            Source = source;
            IsBlocking = isBlocking;

            var properties = watchedProperties
                .Where(property => !string.IsNullOrWhiteSpace(property))
                .ToArray();

            _watchedProperties = properties.Length == 0
                ? null
                : new HashSet<string>(properties, StringComparer.Ordinal);
        }

        public INotifyPropertyChanged Source { get; }
        public Func<bool> IsBlocking { get; }

        public bool ShouldRecompute(string? propertyName)
        {
            if (_watchedProperties == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return true;
            }

            return _watchedProperties.Contains(propertyName);
        }
    }
}
