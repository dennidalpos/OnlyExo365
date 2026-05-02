using System.ComponentModel;
using System.Windows.Input;
using OnlyExo365.Shell.Helpers;
using OnlyExo365.Shell.Services;
using OnlyExo365.Shell.Text;

namespace OnlyExo365.Shell.ViewModels;

public sealed class ShellNavigationStateViewModel : ViewModelBase, IDisposable
{
    private static readonly NavigationPageBinding[] PageBindings =
    [
        new(NavigationPage.Dashboard, nameof(IsDashboardPage)),
        new(NavigationPage.Contacts, nameof(IsContactsPage)),
        new(NavigationPage.Resources, nameof(IsResourcesPage)),
        new(NavigationPage.PublicFolders, nameof(IsPublicFoldersPage)),
        new(NavigationPage.MobileDevices, nameof(IsMobileDevicesPage)),
        new(NavigationPage.Migration, nameof(IsMigrationPage)),
        new(NavigationPage.Permissions, nameof(IsPermissionsPage)),
        new(NavigationPage.Mailboxes, nameof(IsMailboxesPage)),
        new(NavigationPage.DeletedMailboxes, nameof(IsDeletedMailboxesPage)),
        new(NavigationPage.MailboxSpace, nameof(IsMailboxSpacePage)),
        new(NavigationPage.MailboxAccessReport, nameof(IsMailboxAccessReportPage)),
        new(NavigationPage.DistributionLists, nameof(IsDistributionListsPage)),
        new(NavigationPage.MessageTrace, nameof(IsMessageTracePage)),
        new(NavigationPage.Compliance, nameof(IsCompliancePage)),
        new(NavigationPage.MailSecurity, nameof(IsMailSecurityPage)),
        new(NavigationPage.MailFlow, nameof(IsMailFlowPage)),
        new(NavigationPage.Tools, nameof(IsToolsPage)),
        new(NavigationPage.Logs, nameof(IsLogsPage))
    ];

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

        NavigateToDashboardCommand = CreateNavigateCommand(NavigationPage.Dashboard);
        NavigateToContactsCommand = CreateNavigateCommand(NavigationPage.Contacts);
        NavigateToResourcesCommand = CreateNavigateCommand(NavigationPage.Resources);
        NavigateToPublicFoldersCommand = CreateNavigateCommand(NavigationPage.PublicFolders);
        NavigateToMobileDevicesCommand = CreateNavigateCommand(NavigationPage.MobileDevices);
        NavigateToMigrationCommand = CreateNavigateCommand(NavigationPage.Migration);
        NavigateToPermissionsCommand = CreateNavigateCommand(NavigationPage.Permissions);
        NavigateToMailboxesCommand = CreateNavigateCommand(NavigationPage.Mailboxes);
        NavigateToDeletedMailboxesCommand = CreateNavigateCommand(NavigationPage.DeletedMailboxes);
        NavigateToMailboxSpaceCommand = CreateNavigateCommand(NavigationPage.MailboxSpace);
        NavigateToMailboxAccessReportCommand = CreateNavigateCommand(NavigationPage.MailboxAccessReport);
        NavigateToDistributionListsCommand = CreateNavigateCommand(NavigationPage.DistributionLists);
        NavigateToMessageTraceCommand = CreateNavigateCommand(NavigationPage.MessageTrace);
        NavigateToComplianceCommand = CreateNavigateCommand(NavigationPage.Compliance);
        NavigateToMailSecurityCommand = CreateNavigateCommand(NavigationPage.MailSecurity);
        NavigateToMailFlowCommand = CreateNavigateCommand(NavigationPage.MailFlow);
        NavigateToToolsCommand = CreateNavigateCommand(NavigationPage.Tools);
        NavigateToLogsCommand = CreateNavigateCommand(NavigationPage.Logs);
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

    public bool IsDashboardPage => IsCurrentPage(NavigationPage.Dashboard);
    public bool IsContactsPage => IsCurrentPage(NavigationPage.Contacts);
    public bool IsResourcesPage => IsCurrentPage(NavigationPage.Resources);
    public bool IsPublicFoldersPage => IsCurrentPage(NavigationPage.PublicFolders);
    public bool IsMobileDevicesPage => IsCurrentPage(NavigationPage.MobileDevices);
    public bool IsMigrationPage => IsCurrentPage(NavigationPage.Migration);
    public bool IsPermissionsPage => IsCurrentPage(NavigationPage.Permissions);
    public bool IsMailboxesPage => IsCurrentPage(NavigationPage.Mailboxes);
    public bool IsDeletedMailboxesPage => IsCurrentPage(NavigationPage.DeletedMailboxes);
    public bool IsMailboxSpacePage => IsCurrentPage(NavigationPage.MailboxSpace);
    public bool IsMailboxAccessReportPage => IsCurrentPage(NavigationPage.MailboxAccessReport);
    public bool IsDistributionListsPage => IsCurrentPage(NavigationPage.DistributionLists);
    public bool IsMessageTracePage => IsCurrentPage(NavigationPage.MessageTrace);
    public bool IsCompliancePage => IsCurrentPage(NavigationPage.Compliance);
    public bool IsMailSecurityPage => IsCurrentPage(NavigationPage.MailSecurity);
    public bool IsMailFlowPage => IsCurrentPage(NavigationPage.MailFlow);
    public bool IsToolsPage => IsCurrentPage(NavigationPage.Tools);
    public bool IsLogsPage => IsCurrentPage(NavigationPage.Logs);

    public string CurrentPageTitle => UiTextCatalog.GetNavigationLabel(CurrentPage);

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
        foreach (var binding in PageBindings)
        {
            OnPropertyChanged(binding.IsCurrentPagePropertyName);
        }

        OnPropertyChanged(nameof(CurrentPageTitle));
    }

    private ICommand CreateNavigateCommand(NavigationPage page)
        => new RelayCommand(() => _navigationService.NavigateTo(page));

    private bool IsCurrentPage(NavigationPage page)
        => CurrentPage == page;

    private sealed record NavigationPageBinding(NavigationPage Page, string IsCurrentPagePropertyName);

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

