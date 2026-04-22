using System.Text.RegularExpressions;
using System.Windows.Input;
using OnlyExo365.Shell.Services;

namespace OnlyExo365.Shell.ViewModels;

internal static class MailFlowViewModelSupport
{
    private static readonly Regex EmailRegex = new(@"^[^@\s]+@[^@\s]+\.[^@\s]+$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex DomainRegex = new(@"^[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)+$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex RemoteDomainRegex = new(@"^(?:\*\.)?[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)+$", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public static List<string> SplitCsv(string value) => value
        .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
        .Where(v => !string.IsNullOrWhiteSpace(v))
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToList();

    public static bool IsValidEmail(string value) => !string.IsNullOrWhiteSpace(value) && EmailRegex.IsMatch(value.Trim());
    public static bool IsValidDomain(string value) => !string.IsNullOrWhiteSpace(value) && DomainRegex.IsMatch(value.Trim());

    public static bool IsValidRemoteDomain(string value, bool allowDefaultWildcard)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        var trimmed = value.Trim();
        if (trimmed == "*")
        {
            return allowDefaultWildcard;
        }

        return RemoteDomainRegex.IsMatch(trimmed);
    }

    public static bool IsValidOptionalUri(string value)
        => string.IsNullOrWhiteSpace(value) || Uri.TryCreate(value.Trim(), UriKind.Absolute, out _);

    public static bool AreValidDomains(IEnumerable<string> domains) => domains.All(IsValidDomain);
}

internal sealed class MailFlowOperationCoordinator : ViewModelBase
{
    private int _busyCount;
    private string? _errorMessage;
    private string _loadingOverlayText = "Loading Mail Flow workspace...";

    public bool IsLoading => _busyCount > 0;
    public string LoadingOverlayText
    {
        get => _loadingOverlayText;
        private set => SetProperty(ref _loadingOverlayText, value);
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

    public void BeginOperation(string? loadingOverlayText = null)
    {
        var wasLoading = IsLoading;
        if (!string.IsNullOrWhiteSpace(loadingOverlayText))
        {
            LoadingOverlayText = loadingOverlayText;
        }

        _busyCount++;
        if (!wasLoading)
        {
            OnPropertyChanged(nameof(IsLoading));
            CommandManager.InvalidateRequerySuggested();
        }
    }

    public void EndOperation()
    {
        if (_busyCount == 0)
        {
            return;
        }

        _busyCount--;
        if (!IsLoading)
        {
            OnPropertyChanged(nameof(IsLoading));
            CommandManager.InvalidateRequerySuggested();
        }
    }

    public void SetError(string? message) => ErrorMessage = message;
    public void ClearError() => ErrorMessage = null;
}

internal abstract class MailFlowSectionViewModelBase : ViewModelBase
{
    private readonly Func<CancellationToken, Task> _refreshAllAsync;
    private string? _sectionWarningMessage;
    private bool _isSectionSupported = true;

    protected MailFlowSectionViewModelBase(
        IMailFlowWorkerService workerService,
        ShellViewModel shellViewModel,
        MailFlowOperationCoordinator coordinator,
        Func<CancellationToken, Task> refreshAllAsync)
    {
        WorkerService = workerService;
        ShellViewModel = shellViewModel;
        Coordinator = coordinator;
        _refreshAllAsync = refreshAllAsync;
    }

    protected IMailFlowWorkerService WorkerService { get; }
    protected ShellViewModel ShellViewModel { get; }
    protected MailFlowOperationCoordinator Coordinator { get; }
    public string? SectionWarningMessage
    {
        get => _sectionWarningMessage;
        private set
        {
            if (SetProperty(ref _sectionWarningMessage, value))
            {
                OnPropertyChanged(nameof(HasSectionWarning));
            }
        }
    }

    public bool HasSectionWarning => !string.IsNullOrWhiteSpace(SectionWarningMessage);

    public bool IsSectionSupported
    {
        get => _isSectionSupported;
        private set => SetProperty(ref _isSectionSupported, value);
    }

    protected void SetError(string? errorMessage) => Coordinator.SetError(errorMessage);
    protected void InvalidateCommands() => CommandManager.InvalidateRequerySuggested();
    protected void SetSectionState(bool isSupported, string? warningMessage)
    {
        IsSectionSupported = isSupported;
        SectionWarningMessage = warningMessage;
        InvalidateCommands();
    }

    protected async Task RefreshAllAsync(CancellationToken cancellationToken)
    {
        await _refreshAllAsync(cancellationToken);
    }

    protected async Task ExecuteBusyActionAsync(Func<CancellationToken, Task> action, CancellationToken cancellationToken)
    {
        Coordinator.BeginOperation("Applying Mail Flow changes...");
        Coordinator.ClearError();

        try
        {
            await action(cancellationToken);
        }
        catch (Exception ex)
        {
            Coordinator.SetError(ex.Message);
        }
        finally
        {
            Coordinator.EndOperation();
        }
    }

    protected void RaiseProperties(IEnumerable<string> propertyNames)
    {
        foreach (var propertyName in propertyNames)
        {
            OnPropertyChanged(propertyName);
        }
    }
}


