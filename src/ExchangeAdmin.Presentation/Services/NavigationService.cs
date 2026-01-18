namespace ExchangeAdmin.Presentation.Services;

/// <summary>
/// Navigation pages.
/// </summary>
public enum NavigationPage
{
    Dashboard,
    Mailboxes,
    SharedMailboxes,
    DistributionLists,
    Tools,
    Logs
}

/// <summary>
/// Navigation service for page management.
/// </summary>
public class NavigationService
{
    private NavigationPage _currentPage = NavigationPage.Dashboard;
    private string? _selectedIdentity;
    private object? _selectedItem;

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

    public event EventHandler<NavigationPage>? PageChanged;
    public event EventHandler<string?>? SelectedIdentityChanged;
    public event EventHandler<object?>? SelectedItemChanged;
    public event EventHandler<NavigatingEventArgs>? Navigating;

    public void NavigateTo(NavigationPage page)
    {
        // Allow listeners to cancel navigation if there are unsaved changes
        var args = new NavigatingEventArgs(page);
        Navigating?.Invoke(this, args);

        if (args.Cancel)
        {
            return;
        }

        SelectedIdentity = null;
        SelectedItem = null;
        CurrentPage = page;
    }

    public void NavigateToDetails(NavigationPage parentPage, string identity, object? item = null)
    {
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
}

/// <summary>
/// Event args for navigation events that can be cancelled.
/// </summary>
public class NavigatingEventArgs : EventArgs
{
    public NavigatingEventArgs(NavigationPage targetPage)
    {
        TargetPage = targetPage;
    }

    public NavigationPage TargetPage { get; }
    public bool Cancel { get; set; }
}
