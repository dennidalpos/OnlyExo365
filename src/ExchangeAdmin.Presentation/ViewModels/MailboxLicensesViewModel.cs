using System.Globalization;
using System.Collections.ObjectModel;
using System.Windows.Input;
using ExchangeAdmin.Application.Security;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;
using ExchangeAdmin.Presentation.Helpers;

namespace ExchangeAdmin.Presentation.ViewModels;

public sealed class MailboxLicensesViewModel : ViewModelBase
{
    private readonly IMailboxesWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;
    private readonly Func<string?> _getPrimarySmtpAddress;
    private readonly Func<string?> _getDisplayName;
    private readonly Func<CancellationToken, Task>? _afterLicenseMutation;

    private bool _isLicenseLoading;
    private bool _isLicenseSaving;
    private string? _licenseErrorMessage;
    private string? _licenseWriteDisabledMessage;
    private string? _usageLocationSuggestionMessage;
    private string? _usageLocationSuggestedValue;
    private string? _usageLocationSuggestionSource;
    private UsageLocationOption? _selectedUsageLocation;
    private TenantLicenseDto? _selectedLicenseToAdd;

    public MailboxLicensesViewModel(
        IMailboxesWorkerService workerService,
        ShellViewModel shellViewModel,
        Func<string?> getPrimarySmtpAddress,
        Func<string?> getDisplayName,
        Func<CancellationToken, Task>? afterLicenseMutation = null)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;
        _getPrimarySmtpAddress = getPrimarySmtpAddress;
        _getDisplayName = getDisplayName;
        _afterLicenseMutation = afterLicenseMutation;

        _shellViewModel.PropertyChanged += OnShellViewModelPropertyChanged;

        AddLicenseCommand = new AsyncRelayCommand(AddLicenseAsync, () => CanAddLicense && !IsLicenseSaving);
        RemoveLicenseCommand = new AsyncRelayCommand<UserLicenseDto>(RemoveLicenseAsync, license => license != null && CanMutateLicenses && !IsLicenseSaving);
        ApplySuggestedUsageLocationCommand = new AsyncRelayCommand(ApplySuggestedUsageLocationAsync, () => CanApplySuggestedUsageLocation);
        RefreshLicensesCommand = new AsyncRelayCommand(RefreshLicensesAsync, () => !IsLicenseLoading && !string.IsNullOrEmpty(_getPrimarySmtpAddress()));

        RefreshLicenseWriteState();
    }

    public ObservableCollection<UserLicenseDto> AssignedLicenses { get; } = new();
    public ObservableCollection<TenantLicenseDto> AvailableLicenses { get; } = new();
    public ObservableCollection<UsageLocationOption> UsageLocations { get; } = new(CreateUsageLocationOptions());
    public bool HasAssignedLicenses => AssignedLicenses.Count > 0;

    public TenantLicenseDto? SelectedLicenseToAdd
    {
        get => _selectedLicenseToAdd;
        set
        {
            if (SetProperty(ref _selectedLicenseToAdd, value))
            {
                OnPropertyChanged(nameof(CanAddLicense));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool IsLicenseLoading
    {
        get => _isLicenseLoading;
        private set
        {
            if (SetProperty(ref _isLicenseLoading, value))
            {
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool IsLicenseSaving
    {
        get => _isLicenseSaving;
        private set
        {
            if (SetProperty(ref _isLicenseSaving, value))
            {
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public string? LicenseErrorMessage
    {
        get => _licenseErrorMessage;
        private set
        {
            if (SetProperty(ref _licenseErrorMessage, value))
            {
                OnPropertyChanged(nameof(HasLicenseError));
            }
        }
    }

    public bool HasLicenseError => !string.IsNullOrEmpty(LicenseErrorMessage);

    public string? UsageLocationSuggestionMessage
    {
        get => _usageLocationSuggestionMessage;
        private set
        {
            if (SetProperty(ref _usageLocationSuggestionMessage, value))
            {
                OnPropertyChanged(nameof(HasUsageLocationSuggestion));
            }
        }
    }

    public bool HasUsageLocationSuggestion => !string.IsNullOrWhiteSpace(UsageLocationSuggestionMessage);

    public string? UsageLocationSuggestedValue
    {
        get => _usageLocationSuggestedValue;
        private set
        {
            if (SetProperty(ref _usageLocationSuggestedValue, value))
            {
                OnPropertyChanged(nameof(CanApplySuggestedUsageLocation));
                OnPropertyChanged(nameof(UsageLocationActionLabel));
            }
        }
    }

    public string? UsageLocationSuggestionSource
    {
        get => _usageLocationSuggestionSource;
        private set
        {
            if (SetProperty(ref _usageLocationSuggestionSource, value))
            {
                OnPropertyChanged(nameof(UsageLocationActionLabel));
            }
        }
    }

    public UsageLocationOption? SelectedUsageLocation
    {
        get => _selectedUsageLocation;
        set
        {
            if (SetProperty(ref _selectedUsageLocation, value))
            {
                OnPropertyChanged(nameof(CanApplySuggestedUsageLocation));
                OnPropertyChanged(nameof(UsageLocationActionLabel));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public string? LicenseWriteDisabledMessage
    {
        get => _licenseWriteDisabledMessage;
        private set
        {
            if (SetProperty(ref _licenseWriteDisabledMessage, value))
            {
                OnPropertyChanged(nameof(IsLicenseWriteBlocked));
                OnPropertyChanged(nameof(CanMutateLicenses));
                OnPropertyChanged(nameof(LicenseWriteActionTooltip));
                OnPropertyChanged(nameof(CanAddLicense));
            }
        }
    }

    public bool IsLicenseWriteBlocked => !string.IsNullOrWhiteSpace(LicenseWriteDisabledMessage);
    public bool CanMutateLicenses => !IsLicenseWriteBlocked;
    public string LicenseWriteActionTooltip => IsLicenseWriteBlocked
        ? LicenseWriteDisabledMessage!
        : "Assign or remove Microsoft 365 licenses for the selected user.";
    public bool CanAddLicense => SelectedLicenseToAdd != null && !string.IsNullOrEmpty(_getPrimarySmtpAddress()) && CanMutateLicenses;
    public bool CanApplySuggestedUsageLocation => !string.IsNullOrWhiteSpace(GetSelectedUsageLocationCode()) && CanMutateLicenses && !IsLicenseSaving;
    public string UsageLocationActionLabel
    {
        get
        {
            var usageLocationCode = GetSelectedUsageLocationCode();
            if (string.IsNullOrWhiteSpace(usageLocationCode))
            {
                return "Refresh UsageLocation";
            }

            return SelectedLicenseToAdd != null
                ? $"Apply {usageLocationCode} and retry"
                : $"Set UsageLocation to {usageLocationCode}";
        }
    }

    public ICommand AddLicenseCommand { get; }
    public ICommand RemoveLicenseCommand { get; }
    public ICommand ApplySuggestedUsageLocationCommand { get; }
    public ICommand RefreshLicensesCommand { get; }

    public void LoadAssignedLicenses(IEnumerable<UserLicenseDto>? licenses)
    {
        AssignedLicenses.Clear();
        if (licenses != null)
        {
            foreach (var license in licenses)
            {
                AssignedLicenses.Add(license);
            }
        }

        OnPropertyChanged(nameof(HasAssignedLicenses));
    }

    public void Reset()
    {
        AssignedLicenses.Clear();
        AvailableLicenses.Clear();
        SelectedUsageLocation = null;
        SelectedLicenseToAdd = null;
        LicenseErrorMessage = null;
        ClearUsageLocationSuggestion();
        RefreshLicenseWriteState();
        OnPropertyChanged(nameof(HasAssignedLicenses));
    }

    public Task RefreshAsync(CancellationToken cancellationToken = default)
        => RefreshLicensesAsync(cancellationToken);

    public async Task LoadAvailableLicensesAsync(bool isExchangeConnected, CancellationToken cancellationToken)
    {
        if (!isExchangeConnected)
        {
            return;
        }

        RefreshLicenseWriteState();
        IsLicenseLoading = true;
        try
        {
            var result = await _workerService.GetAvailableLicensesAsync(cancellationToken: cancellationToken);
            if (result.IsSuccess && result.Value != null)
            {
                RunOnUiThread(() =>
                {
                    AvailableLicenses.Clear();
                    foreach (var license in result.Value.Licenses.OrderBy(item => item.SkuPartNumber))
                    {
                        AvailableLicenses.Add(license);
                    }
                });

                await _shellViewModel.RefreshConnectionStatusAsync(cancellationToken);
            }
            else if (!result.WasCancelled)
            {
                _shellViewModel.AddLog(LogLevel.Warning, result.Error?.Message ?? "Unable to retrieve available licenses.");
            }
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Error, $"Available licenses load failed: {ex.Message}");
        }
        finally
        {
            IsLicenseLoading = false;
        }
    }

    private async Task RefreshLicensesAsync(CancellationToken cancellationToken)
    {
        var primarySmtpAddress = _getPrimarySmtpAddress();
        if (string.IsNullOrEmpty(primarySmtpAddress))
        {
            return;
        }

        IsLicenseLoading = true;
        LicenseErrorMessage = null;
        ClearUsageLocationSuggestion();
        RefreshLicenseWriteState();

        try
        {
            var result = await _workerService.GetUserLicensesAsync(
                new GetUserLicensesRequest { UserPrincipalName = primarySmtpAddress },
                cancellationToken: cancellationToken);

            if (result.IsSuccess && result.Value != null)
            {
                LoadAssignedLicenses(result.Value.Licenses);
                await _shellViewModel.RefreshConnectionStatusAsync(cancellationToken);
            }
            else if (!result.WasCancelled)
            {
                LicenseErrorMessage = result.Error?.Message ?? "Unable to retrieve user licenses.";
            }

            await LoadAvailableLicensesAsync(true, cancellationToken);
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            LicenseErrorMessage = ex.Message;
        }
        finally
        {
            IsLicenseLoading = false;
        }
    }

    private async Task AddLicenseAsync(CancellationToken cancellationToken)
    {
        var primarySmtpAddress = _getPrimarySmtpAddress();
        if (SelectedLicenseToAdd == null || string.IsNullOrEmpty(primarySmtpAddress))
        {
            return;
        }

        if (!EnsureLicenseWritesAllowed())
        {
            return;
        }

        IsLicenseSaving = true;
        LicenseErrorMessage = null;
        ClearUsageLocationSuggestion();

        try
        {
            var result = await ExecuteLicenseMutationAsync(
                new SetUserLicenseRequest
                {
                    UserPrincipalName = primarySmtpAddress,
                    AddLicenseSkuIds = new List<string> { SelectedLicenseToAdd.SkuId }
                },
                $"Assigning license {SelectedLicenseToAdd.SkuPartNumber}...",
                $"License {SelectedLicenseToAdd.SkuPartNumber} assigned",
                "Unable to assign the license.",
                clearSelectedLicenseOnSuccess: true,
                cancellationToken: cancellationToken);
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            LicenseErrorMessage = NormalizeLicenseErrorMessage(ex.Message);
            _shellViewModel.AddLog(LogLevel.Error, $"Add license error: {ex.Message}");
        }
        finally
        {
            IsLicenseSaving = false;
        }
    }

    private async Task RemoveLicenseAsync(UserLicenseDto? license, CancellationToken cancellationToken)
    {
        var primarySmtpAddress = _getPrimarySmtpAddress();
        if (license == null || string.IsNullOrEmpty(primarySmtpAddress))
        {
            return;
        }

        if (!EnsureLicenseWritesAllowed())
        {
            return;
        }

        IsLicenseSaving = true;
        LicenseErrorMessage = null;
        ClearUsageLocationSuggestion();

        try
        {
            await ExecuteLicenseMutationAsync(
                new SetUserLicenseRequest
                {
                    UserPrincipalName = primarySmtpAddress,
                    RemoveLicenseSkuIds = new List<string> { license.SkuId }
                },
                $"Removing license {license.SkuPartNumber}...",
                $"License {license.SkuPartNumber} removed",
                "Unable to remove the license.",
                clearSelectedLicenseOnSuccess: false,
                cancellationToken: cancellationToken);
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            LicenseErrorMessage = NormalizeLicenseErrorMessage(ex.Message);
            _shellViewModel.AddLog(LogLevel.Error, $"Remove license error: {ex.Message}");
        }
        finally
        {
            IsLicenseSaving = false;
        }
    }

    private async Task ApplySuggestedUsageLocationAsync(CancellationToken cancellationToken)
    {
        var primarySmtpAddress = _getPrimarySmtpAddress();
        var usageLocationToApply = GetSelectedUsageLocationCode();
        if (string.IsNullOrWhiteSpace(primarySmtpAddress) || string.IsNullOrWhiteSpace(usageLocationToApply))
        {
            return;
        }

        if (!EnsureLicenseWritesAllowed())
        {
            return;
        }

        IsLicenseSaving = true;
        LicenseErrorMessage = null;

        try
        {
            _shellViewModel.AddLog(LogLevel.Information, $"Updating UsageLocation to {usageLocationToApply}...");

            var result = await _workerService.SetUserUsageLocationAsync(
                new SetUserUsageLocationRequest
                {
                    UserPrincipalName = primarySmtpAddress,
                    UsageLocation = usageLocationToApply
                },
                cancellationToken: cancellationToken);

            if (!result.IsSuccess)
            {
                if (!result.WasCancelled)
                {
                    LicenseErrorMessage = NormalizeLicenseErrorMessage(result.Error?.Message ?? "Unable to update UsageLocation.");
                }

                return;
            }

            _shellViewModel.AddLog(LogLevel.Information, $"UsageLocation updated to {usageLocationToApply}");
            ClearUsageLocationSuggestion();

            if (SelectedLicenseToAdd != null)
            {
                await ExecuteLicenseMutationAsync(
                    new SetUserLicenseRequest
                    {
                        UserPrincipalName = primarySmtpAddress,
                        AddLicenseSkuIds = new List<string> { SelectedLicenseToAdd.SkuId }
                    },
                    $"Assigning license {SelectedLicenseToAdd.SkuPartNumber}...",
                    $"License {SelectedLicenseToAdd.SkuPartNumber} assigned",
                    "Unable to assign the license.",
                    clearSelectedLicenseOnSuccess: true,
                    cancellationToken: cancellationToken);
            }
            else
            {
                await _shellViewModel.RefreshConnectionStatusAsync(cancellationToken);
                await RefreshLicensesAsync(cancellationToken);
            }
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            LicenseErrorMessage = NormalizeLicenseErrorMessage(ex.Message);
            _shellViewModel.AddLog(LogLevel.Error, $"Set usage location error: {ex.Message}");
        }
        finally
        {
            IsLicenseSaving = false;
        }
    }

    private async Task<Result> ExecuteLicenseMutationAsync(
        SetUserLicenseRequest request,
        string startLogMessage,
        string successLogMessage,
        string failureMessage,
        bool clearSelectedLicenseOnSuccess,
        CancellationToken cancellationToken)
    {
        _shellViewModel.AddLog(LogLevel.Information, startLogMessage);

        var result = await _workerService.SetUserLicenseAsync(request, cancellationToken: cancellationToken);

        if (result.IsSuccess)
        {
            _shellViewModel.AddLog(LogLevel.Information, successLogMessage);
            await _shellViewModel.RefreshConnectionStatusAsync(cancellationToken);
            if (clearSelectedLicenseOnSuccess)
            {
                SelectedLicenseToAdd = null;
            }

            await RefreshLicensesAsync(cancellationToken);
            if (_afterLicenseMutation != null)
            {
                await _afterLicenseMutation(cancellationToken);
            }

            return result;
        }

        if (!result.WasCancelled)
        {
            LicenseErrorMessage = NormalizeLicenseErrorMessage(result.Error?.Message ?? failureMessage);
            await TryLoadUsageLocationSuggestionAsync(cancellationToken);
        }

        return result;
    }

    private bool EnsureLicenseWritesAllowed()
    {
        RefreshLicenseWriteState();
        if (!IsLicenseWriteBlocked)
        {
            return true;
        }

        LicenseErrorMessage = LicenseWriteDisabledMessage;
        return false;
    }

    private void RefreshLicenseWriteState()
    {
        var evaluation = _shellViewModel.EvaluateLeastPrivilege(LeastPrivilegeCatalog.MailboxLicensingWrite);
        LicenseWriteDisabledMessage = evaluation.Status == LeastPrivilegeFeatureStatus.Blocked
            ? BuildLicenseWriteDisabledMessage(evaluation)
            : null;

        CommandManager.InvalidateRequerySuggested();
    }

    private static string BuildLicenseWriteDisabledMessage(LeastPrivilegeFeatureEvaluation evaluation)
    {
        const string scopeName = "LicenseAssignment.ReadWrite.All";
        var baseMessage =
            $"License changes are disabled by the current configuration. Add {scopeName} to graphLicenseWriteScopes or EXCHANGEADMIN_GRAPH_LICENSE_WRITE_SCOPES and reconnect the session.";

        return evaluation.HasMissingRequirements
            ? $"{baseMessage} Validazione: {evaluation.MissingRequirementsDisplay}"
            : baseMessage;
    }

    internal static string NormalizeLicenseErrorMessage(string? message)
    {
        return string.IsNullOrWhiteSpace(message)
            ? "Unable to complete the license operation."
            : message.Trim();
    }

    private async Task TryLoadUsageLocationSuggestionAsync(CancellationToken cancellationToken)
    {
        var primarySmtpAddress = _getPrimarySmtpAddress();
        if (string.IsNullOrWhiteSpace(primarySmtpAddress) ||
            !(LicenseErrorMessage?.Contains("usage location", StringComparison.OrdinalIgnoreCase) ?? false) &&
            !(LicenseErrorMessage?.Contains("UsageLocation", StringComparison.OrdinalIgnoreCase) ?? false))
        {
            return;
        }

        var result = await _workerService.GetUsageLocationSuggestionAsync(
            new GetUsageLocationSuggestionRequest
            {
                UserPrincipalName = primarySmtpAddress
            },
            cancellationToken: cancellationToken);

        if (!result.IsSuccess || result.Value == null || string.IsNullOrWhiteSpace(result.Value.SuggestedUsageLocation))
        {
            return;
        }

        UsageLocationSuggestedValue = result.Value.SuggestedUsageLocation;
        UsageLocationSuggestionSource = result.Value.SuggestionSource;
        UsageLocationSuggestionMessage = BuildUsageLocationSuggestionMessage(result.Value);
        SelectUsageLocation(result.Value.SuggestedUsageLocation);
    }

    internal static string BuildUsageLocationSuggestionMessage(GetUsageLocationSuggestionResponse suggestion)
    {
        var sourceLabel = suggestion.SuggestionSource switch
        {
            "User" => "user",
            "Tenant" => "tenant",
            "Configuration" => "local configuration",
            _ => "unspecified source"
        };

        var details = string.IsNullOrWhiteSpace(suggestion.SuggestionDetails)
            ? string.Empty
            : $" {suggestion.SuggestionDetails}";

        return $"Suggested UsageLocation: {suggestion.SuggestedUsageLocation} (source: {sourceLabel}).{details}";
    }

    private void ClearUsageLocationSuggestion()
    {
        UsageLocationSuggestionMessage = null;
        UsageLocationSuggestedValue = null;
        UsageLocationSuggestionSource = null;
    }

    private string? GetSelectedUsageLocationCode()
        => SelectedUsageLocation?.Code ?? UsageLocationSuggestedValue;

    private void SelectUsageLocation(string? usageLocationCode)
    {
        if (string.IsNullOrWhiteSpace(usageLocationCode))
        {
            return;
        }

        var normalizedCode = ExchangeOnlineConfiguration.NormalizeUsageLocation(usageLocationCode);
        var match = UsageLocations.FirstOrDefault(option =>
            string.Equals(option.Code, normalizedCode, StringComparison.OrdinalIgnoreCase));

        if (match != null)
        {
            SelectedUsageLocation = match;
        }
    }

    internal static IReadOnlyList<UsageLocationOption> CreateUsageLocationOptions()
    {
        return CultureInfo
            .GetCultures(CultureTypes.SpecificCultures)
            .Select(static culture =>
            {
                try
                {
                    return new RegionInfo(culture.Name);
                }
                catch (ArgumentException)
                {
                    return null;
                }
            })
            .Where(static region => region != null && region.TwoLetterISORegionName.Length == 2 && region.TwoLetterISORegionName.All(char.IsLetter))
            .GroupBy(static region => region!.TwoLetterISORegionName.ToUpperInvariant(), StringComparer.OrdinalIgnoreCase)
            .Select(static group =>
            {
                var displayName = group
                    .Select(static region => region!.EnglishName.Trim())
                    .Where(static name => !string.IsNullOrWhiteSpace(name))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(static name => name, StringComparer.OrdinalIgnoreCase)
                    .First();

                return new UsageLocationOption(group.Key, displayName);
            })
            .OrderBy(static option => option.DisplayName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(static option => option.Code, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private void OnShellViewModelPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (e.PropertyName is nameof(ShellViewModel.ExchangeState) or nameof(ShellViewModel.Capabilities))
        {
            RefreshLicenseWriteState();
        }
    }
}

public sealed class UsageLocationOption
{
    public UsageLocationOption(string code, string displayName)
    {
        Code = code;
        DisplayName = displayName;
    }

    public string Code { get; }
    public string DisplayName { get; }
    public string DisplayLabel => $"{DisplayName} ({Code})";
}
