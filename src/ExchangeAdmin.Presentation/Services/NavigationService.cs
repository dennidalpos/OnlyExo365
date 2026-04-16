namespace ExchangeAdmin.Presentation.Services;

public enum NavigationPage
{
    Dashboard,
    Contacts,
    Resources,
    PublicFolders,
    MobileDevices,
    Migration,
    Permissions,
    Mailboxes,
    DeletedMailboxes,
    MailboxSpace,
    MailboxAccessReport,
    DistributionLists,
    MessageTrace,
    Compliance,
    MailSecurity,
    MailFlow,
    Tools,
    Logs
}

public class NavigationService
{
    private NavigationPage _currentPage = NavigationPage.Dashboard;
    private string? _selectedIdentity;
    private object? _selectedItem;
    private bool _isNavigationPending;
    private NavigationPage? _pendingPage;

    public NavigationPage CurrentPage
    {
        get => _currentPage;
        private set
        {
            if (_currentPage != value)
            {
                _currentPage = value;
                PageChanged?.Invoke(this, value);
            }
        }
    }

    public string? SelectedIdentity
    {
        get => _selectedIdentity;
        private set
        {
            _selectedIdentity = value;
            SelectedIdentityChanged?.Invoke(this, value);
        }
    }

    public object? SelectedItem
    {
        get => _selectedItem;
        private set
        {
            _selectedItem = value;
            SelectedItemChanged?.Invoke(this, value);
        }
    }

    public bool IsNavigationPending
    {
        get => _isNavigationPending;
        private set
        {
            if (_isNavigationPending == value)
            {
                return;
            }

            _isNavigationPending = value;
            NavigationStateChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    public NavigationPage? PendingPage => _pendingPage;

    public event EventHandler<NavigationPage>? PageChanged;
    public event EventHandler<string?>? SelectedIdentityChanged;
    public event EventHandler<object?>? SelectedItemChanged;
    public event EventHandler<NavigatingEventArgs>? Navigating;
    public event EventHandler? NavigationStateChanged;

    public void NavigateTo(NavigationPage page)
    {
        var args = new NavigatingEventArgs(page);
        Navigating?.Invoke(this, args);

        if (args.Cancel)
        {
            return;
        }

        var requiresPageTransition = CurrentPage != page;
        if (requiresPageTransition)
        {
            BeginNavigation(page);
        }

        SelectedIdentity = null;
        SelectedItem = null;
        CurrentPage = page;
    }

    public void NavigateToDetails(NavigationPage parentPage, string identity, object? item = null)
    {
        var args = new NavigatingEventArgs(parentPage);
        Navigating?.Invoke(this, args);

        if (args.Cancel)
        {
            return;
        }

        if (CurrentPage != parentPage)
        {
            BeginNavigation(parentPage);
        }

        CurrentPage = parentPage;
        SelectedIdentity = identity;
        SelectedItem = item;
    }

    public void ClearSelection()
    {
        SelectedIdentity = null;
        SelectedItem = null;
    }

    public bool HasSelection => !string.IsNullOrEmpty(SelectedIdentity);

    public void CompleteNavigation(NavigationPage page)
    {
        if (!_pendingPage.HasValue || _pendingPage.Value != page)
        {
            return;
        }

        _pendingPage = null;
        IsNavigationPending = false;
    }

    private void BeginNavigation(NavigationPage page)
    {
        _pendingPage = page;
        IsNavigationPending = true;
    }
}

public class NavigatingEventArgs : EventArgs
{
    public NavigatingEventArgs(NavigationPage targetPage)
    {
        TargetPage = targetPage;
    }

    public NavigationPage TargetPage { get; }
    public bool Cancel { get; set; }
}
