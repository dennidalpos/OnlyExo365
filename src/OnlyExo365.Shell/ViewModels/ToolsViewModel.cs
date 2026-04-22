using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.ComponentModel;
using System.Windows.Input;
using OnlyExo365.Shell.Security;
using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Shell.Helpers;
using OnlyExo365.Shell.Localization;

namespace OnlyExo365.Shell.ViewModels;

public class ToolsViewModel : ViewModelBase, IDisposable
{
    private readonly ISystemWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;
    private readonly LeastPrivilegeEvaluator _leastPrivilegeEvaluator;
    private readonly ObservableCollection<LeastPrivilegeFeatureDisplay> _leastPrivilegeMatrix = [];

    private bool _isChecking;
    private bool _isInstalling;
    private string? _errorMessage;
    private double _installProgress;
    private string? _installStatus;

    private string? _powerShellVersion;
    private bool _isPowerShell7;
    private bool _exchangeModuleInstalled;
    private string? _exchangeModuleVersion;
    private bool _graphModuleInstalled;
    private string? _graphModuleVersion;
    private bool _hasChecked;
    private string _prerequisiteSummary = string.Empty;
    private string _suggestedActions = string.Empty;
    private string _leastPrivilegeSummary = string.Empty;
    private PrerequisiteStatusDto? _lastPrerequisiteStatus;

    public ToolsViewModel(
        ISystemWorkerService workerService,
        ShellViewModel shellViewModel,
        ExchangeOnlineConfiguration? exchangeConfiguration = null,
        LicenseCatalogViewModel? catalogViewModel = null)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;
        _leastPrivilegeEvaluator = new LeastPrivilegeEvaluator(exchangeConfiguration ?? ExchangeOnlineConfiguration.CreateDefault());

        Catalog = catalogViewModel;

        CheckPrerequisitesCommand = new AsyncRelayCommand(CheckPrerequisitesAsync, () => !IsChecking && !IsInstalling);
        InstallPowerShellCommand = new AsyncRelayCommand(InstallPowerShellAsync, () => !IsInstalling && !IsPowerShell7);
        InstallExchangeModuleCommand = new AsyncRelayCommand(() => InstallModuleAsync("ExchangeOnlineManagement"), () => !IsInstalling);
        InstallGraphModuleCommand = new AsyncRelayCommand(() => InstallModuleAsync("Microsoft.Graph"), () => !IsInstalling);
        OpenPowerShellGitHubCommand = new RelayCommand(OpenPowerShellGitHub);

        _shellViewModel.PropertyChanged += OnShellPropertyChanged;
        LocalizationService.Instance.CultureChanged += OnCultureChanged;
        RefreshLeastPrivilegeMatrix();
    }

    public bool IsChecking
    {
        get => _isChecking;
        private set
        {
            if (SetProperty(ref _isChecking, value))
            {
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool IsInstalling
    {
        get => _isInstalling;
        private set
        {
            if (SetProperty(ref _isInstalling, value))
            {
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

    public bool HasError => !string.IsNullOrEmpty(ErrorMessage);

    public double InstallProgress
    {
        get => _installProgress;
        private set => SetProperty(ref _installProgress, value);
    }

    public string? InstallStatus
    {
        get => _installStatus;
        private set => SetProperty(ref _installStatus, value);
    }

    public string? PowerShellVersion
    {
        get => _powerShellVersion;
        private set => SetProperty(ref _powerShellVersion, value);
    }

    public bool IsPowerShell7
    {
        get => _isPowerShell7;
        private set
        {
            if (SetProperty(ref _isPowerShell7, value))
            {
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool ExchangeModuleInstalled
    {
        get => _exchangeModuleInstalled;
        private set => SetProperty(ref _exchangeModuleInstalled, value);
    }

    public string? ExchangeModuleVersion
    {
        get => _exchangeModuleVersion;
        private set => SetProperty(ref _exchangeModuleVersion, value);
    }

    public bool GraphModuleInstalled
    {
        get => _graphModuleInstalled;
        private set => SetProperty(ref _graphModuleInstalled, value);
    }

    public string? GraphModuleVersion
    {
        get => _graphModuleVersion;
        private set => SetProperty(ref _graphModuleVersion, value);
    }

    public bool HasChecked
    {
        get => _hasChecked;
        private set => SetProperty(ref _hasChecked, value);
    }

    public LicenseCatalogViewModel? Catalog { get; }

    public ICommand CheckPrerequisitesCommand { get; }
    public ICommand InstallPowerShellCommand { get; }
    public ICommand InstallExchangeModuleCommand { get; }
    public ICommand InstallGraphModuleCommand { get; }
    public ICommand OpenPowerShellGitHubCommand { get; }
    public ICommand SetWorkerConsoleVisibilityCommand => _shellViewModel.SetWorkerConsoleVisibilityCommand;

    public ObservableCollection<LeastPrivilegeFeatureDisplay> LeastPrivilegeMatrix => _leastPrivilegeMatrix;

    public bool HasLeastPrivilegeMatrix => LeastPrivilegeMatrix.Count > 0;

    public string PrerequisiteSummary
    {
        get => _prerequisiteSummary;
        private set => SetProperty(ref _prerequisiteSummary, value);
    }

    public string SuggestedActions
    {
        get => _suggestedActions;
        private set => SetProperty(ref _suggestedActions, value);
    }

    public string LeastPrivilegeSummary
    {
        get => _leastPrivilegeSummary;
        private set => SetProperty(ref _leastPrivilegeSummary, value);
    }

    public async Task LoadAsync()
    {
        RefreshLeastPrivilegeMatrix();

        if (!HasChecked && _shellViewModel.IsWorkerRunning)
        {
            await CheckPrerequisitesAsync(CancellationToken.None);
        }
    }

    private async Task CheckPrerequisitesAsync(CancellationToken cancellationToken)
    {
        IsChecking = true;
        ErrorMessage = null;
        InstallProgress = 0;
        InstallStatus = Loc.Get("Tools.Prerequisites.Checking");

        try
        {
            InstallProgress = 20;
            var result = await _workerService.CheckPrerequisitesAsync(
                cancellationToken: cancellationToken);

            InstallProgress = 80;

            if (result.IsSuccess && result.Value != null)
            {
                var status = result.Value;
                PowerShellVersion = status.PowerShellVersion;
                IsPowerShell7 = status.IsPowerShell7;
                ExchangeModuleInstalled = status.ExchangeModuleInstalled;
                ExchangeModuleVersion = status.ExchangeModuleVersion;
                GraphModuleInstalled = status.GraphModuleInstalled;
                GraphModuleVersion = status.GraphModuleVersion;
                HasChecked = true;
                _lastPrerequisiteStatus = status;

                _shellViewModel.AddLog(LogLevel.Information,
                    $"[Prerequisites] PS={status.PowerShellVersion} (pwsh7={status.IsPowerShell7}, policy={status.CurrentUserExecutionPolicy ?? "Unknown"}), EXO={status.ExchangeModuleInstalled}, Graph={status.GraphModuleInstalled}");

                RefreshPrerequisiteText();
            }
            else if (!result.WasCancelled)
            {
                ErrorMessage = result.Error?.Message ?? Loc.Get("Tools.Prerequisites.CheckFailed");
                _shellViewModel.AddLog(LogLevel.Error, $"[Prerequisites] Check failed: {ErrorMessage}");
            }

            InstallProgress = 100;
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
        finally
        {
            IsChecking = false;
            InstallProgress = 0;
            InstallStatus = null;
        }
    }

    private async Task InstallPowerShellAsync(CancellationToken cancellationToken)
    {
        IsInstalling = true;
        ErrorMessage = null;
        InstallProgress = 0;
        InstallStatus = Loc.Get("Tools.Prerequisites.InstallingPowerShell");

        try
        {
            InstallProgress = 20;
            var result = await _workerService.InstallModuleAsync(
                new InstallModuleRequest
                {
                    ModuleName = "PowerShell7",
                    InstallTarget = "PowerShell7",
                    PackageId = "Microsoft.PowerShell"
                },
                cancellationToken: cancellationToken);

            InstallProgress = 90;

            if (result.IsSuccess && result.Value != null)
            {
                if (result.Value.Success)
                {
                    _shellViewModel.AddLog(LogLevel.Information, "[ModuleInstall] PowerShell 7 installed through winget");
                    InstallStatus = Loc.Get("Tools.Prerequisites.PowerShellInstalled");
                }
                else
                {
                    ErrorMessage = BuildInstallErrorMessage(
                        result.Value.Message,
                        result.Value.ManualInstructions,
                        "Download it from https://github.com/PowerShell/PowerShell/releases");
                    _shellViewModel.AddLog(LogLevel.Warning, $"[ModuleInstall] PowerShell 7 install failed: {result.Value.Message}");
                }
            }
            else if (!result.WasCancelled)
            {
                ErrorMessage = result.Error?.Message ?? Loc.Get("Tools.Prerequisites.InstallError");
            }

            InstallProgress = 100;
        }
        catch (Exception ex)
        {
            ErrorMessage = BuildInstallErrorMessage(
                $"Error: {ex.Message}",
                Loc.Get("Tools.Prerequisites.PowerShellManualInstructions"),
                "https://github.com/PowerShell/PowerShell/releases");
        }
        finally
        {
            IsInstalling = false;
            InstallProgress = 0;
            InstallStatus = null;
        }
    }

    private async Task InstallModuleAsync(string moduleName)
    {
        IsInstalling = true;
        ErrorMessage = null;
        InstallProgress = 0;
        InstallStatus = Loc.GetFormat("Tools.Prerequisites.InstallingModule", moduleName);

        try
        {
            InstallProgress = 10;
            var result = await _workerService.InstallModuleAsync(
                new InstallModuleRequest { ModuleName = moduleName },
                cancellationToken: CancellationToken.None);

            InstallProgress = 90;

            if (result.IsSuccess && result.Value != null)
            {
                if (result.Value.Success)
                {
                    var installedVersion = string.IsNullOrWhiteSpace(result.Value.InstalledVersion)
                        ? Loc.Get("Tools.Prerequisites.ApprovedVersion")
                        : Loc.GetFormat("Tools.Prerequisites.VersionFormat", result.Value.InstalledVersion);
                    _shellViewModel.AddLog(LogLevel.Information, $"[ModuleInstall] {moduleName} installed successfully ({installedVersion})");
                    InstallStatus = Loc.GetFormat("Tools.Prerequisites.ModuleInstalled", moduleName, installedVersion);
                    await CheckPrerequisitesAsync(CancellationToken.None);
                }
                else
                {
                    ErrorMessage = BuildInstallErrorMessage(
                        $"Installation failed for {moduleName}: {result.Value.Message}",
                        result.Value.ManualInstructions,
                        $"Install-Module {moduleName} -Force -AllowClobber -Scope CurrentUser");
                    _shellViewModel.AddLog(LogLevel.Error, $"[ModuleInstall] {moduleName} install failed: {result.Value.Message}");
                }
            }
            else if (!result.WasCancelled)
            {
                ErrorMessage = result.Error?.Message ?? Loc.GetFormat("Tools.Prerequisites.ModuleInstallFailed", moduleName);
            }

            InstallProgress = 100;
        }
        catch (Exception ex)
        {
            ErrorMessage = BuildInstallErrorMessage(
                $"Error: {ex.Message}",
                null,
                $"Install-Module {moduleName} -Repository PSGallery -RequiredVersion <approved-version> -Force -AllowClobber -Scope CurrentUser");
        }
        finally
        {
            IsInstalling = false;
            InstallProgress = 0;
            InstallStatus = null;
        }
    }

    private void OpenPowerShellGitHub()
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "https://github.com/PowerShell/PowerShell/releases",
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Error, $"Unable to open the browser: {ex.Message}");
        }
    }

    private void OnShellPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (string.Equals(e.PropertyName, nameof(ShellViewModel.Capabilities), StringComparison.Ordinal) ||
            string.Equals(e.PropertyName, nameof(ShellViewModel.IsExchangeConnected), StringComparison.Ordinal))
        {
            RefreshLeastPrivilegeMatrix();
        }
    }

    private void OnCultureChanged(object? sender, EventArgs e)
    {
        RefreshPrerequisiteText();
        RefreshLeastPrivilegeMatrix();
    }

    private void RefreshLeastPrivilegeMatrix()
    {
        var evaluations = _leastPrivilegeEvaluator.EvaluateAll(_shellViewModel.Capabilities);

        _leastPrivilegeMatrix.Clear();
        foreach (var evaluation in evaluations)
        {
            _leastPrivilegeMatrix.Add(new LeastPrivilegeFeatureDisplay(evaluation));
        }

        OnPropertyChanged(nameof(HasLeastPrivilegeMatrix));

        var readyCount = evaluations.Count(item => item.Status == LeastPrivilegeFeatureStatus.Available);
        var blockedCount = evaluations.Count(item => item.Status == LeastPrivilegeFeatureStatus.Blocked);
        var pendingCount = evaluations.Count(item => item.Status != LeastPrivilegeFeatureStatus.Available &&
                                                     item.Status != LeastPrivilegeFeatureStatus.Blocked);

        LeastPrivilegeSummary = Loc.GetFormat("Tools.LeastPrivilege.Summary", readyCount, blockedCount, pendingCount);
    }

    private void RefreshPrerequisiteText()
    {
        if (_lastPrerequisiteStatus == null)
        {
            return;
        }

        var status = _lastPrerequisiteStatus;
        var exchangeModuleState = BuildModuleState(
            status.ExchangeModuleInstalled,
            status.IsExchangeModuleApproved,
            status.ExchangeModuleVersion,
            status.ExchangeModuleRequiredVersion);
        var graphModuleState = BuildModuleState(
            status.GraphModuleInstalled,
            status.IsGraphModuleApproved,
            status.GraphModuleVersion,
            status.GraphModuleRequiredVersion);

        PrerequisiteSummary = Loc.GetFormat(
            "Tools.Prerequisites.Summary",
            status.PowerShellVersion ?? Loc.Get("Status.Unknown"),
            status.IsExecutionPolicyCompatible ? Loc.Get("Common.OK") : (status.CurrentUserExecutionPolicy ?? Loc.Get("Status.Unknown")),
            exchangeModuleState,
            graphModuleState);

        var suggestions = new List<string>();
        if (!status.IsPowerShell7) suggestions.Add(Loc.Get("Tools.Prerequisites.SuggestInstallPowerShell"));
        if (!status.IsExecutionPolicyCompatible) suggestions.Add(status.ManualInstructions ?? Loc.Get("Tools.Prerequisites.SuggestExecutionPolicy"));
        if (!status.ExchangeModuleInstalled) suggestions.Add(Loc.GetFormat("Tools.Prerequisites.SuggestInstallExchangeModule", status.ExchangeModuleRequiredVersion ?? Loc.Get("Status.Unknown")));
        else if (!status.IsExchangeModuleApproved) suggestions.Add(Loc.GetFormat("Tools.Prerequisites.SuggestAlignExchangeModule", status.ExchangeModuleRequiredVersion ?? Loc.Get("Status.Unknown")));
        if (!status.GraphModuleInstalled) suggestions.Add(Loc.GetFormat("Tools.Prerequisites.SuggestInstallGraphModule", status.GraphModuleRequiredVersion ?? Loc.Get("Status.Unknown")));
        else if (!status.IsGraphModuleApproved) suggestions.Add(Loc.GetFormat("Tools.Prerequisites.SuggestAlignGraphModule", status.GraphModuleRequiredVersion ?? Loc.Get("Status.Unknown")));
        SuggestedActions = suggestions.Count == 0 ? Loc.Get("Tools.Prerequisites.AllSatisfied") : string.Join(" ", suggestions);
    }

    private static string BuildModuleState(bool installed, bool approved, string? version, string? requiredVersion)
    {
        if (!installed)
        {
            return Loc.Get("Status.Module.Missing");
        }

        return approved
            ? Loc.GetFormat("Status.Module.OKVersion", version ?? Loc.Get("Status.Unknown"))
            : Loc.GetFormat("Status.Module.Drift", version ?? Loc.Get("Status.Unknown"), requiredVersion ?? Loc.Get("Status.Unknown"));
    }

    private static string BuildInstallErrorMessage(string header, string? manualInstructions, string fallbackInstructions)
    {
        var instructions = string.IsNullOrWhiteSpace(manualInstructions)
            ? fallbackInstructions
            : manualInstructions;

        return Loc.GetFormat("Tools.Prerequisites.ManualInstallation", header, instructions);
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (disposing)
        {
            _shellViewModel.PropertyChanged -= OnShellPropertyChanged;
            LocalizationService.Instance.CultureChanged -= OnCultureChanged;
            Catalog?.Dispose();
        }
    }
}

public sealed class LeastPrivilegeFeatureDisplay
{
    private readonly LeastPrivilegeFeatureEvaluation _evaluation;

    public LeastPrivilegeFeatureDisplay(LeastPrivilegeFeatureEvaluation evaluation)
    {
        _evaluation = evaluation;
    }

    public string FeatureName => GetFeatureText("Name", _evaluation.FeatureName);

    public string ModuleName => _evaluation.ModuleName switch
    {
        "Dashboard / Mailboxes" => Loc.Get("Tools.LeastPrivilege.Module.DashboardMailboxes"),
        "Mail Flow" => Loc.Get("Nav.MailFlow"),
        "Mail Security" => Loc.Get("Nav.MailSecurity"),
        "Message Trace" => Loc.Get("Nav.MessageTrace"),
        "Mobile Devices" => Loc.Get("Nav.MobileDevices"),
        "Permissions" => Loc.Get("Nav.RoleGroups"),
        _ => _evaluation.ModuleName
    };

    public string Description => GetFeatureText("Description", _evaluation.Description);

    public string StatusLabel => _evaluation.Status switch
    {
        LeastPrivilegeFeatureStatus.Available => Loc.Get("Tools.LeastPrivilege.Status.Ready"),
        LeastPrivilegeFeatureStatus.NeedsAdditionalSession => Loc.Get("Tools.LeastPrivilege.Status.NeedsSession"),
        LeastPrivilegeFeatureStatus.Blocked => Loc.Get("Tools.LeastPrivilege.Status.Blocked"),
        _ => Loc.Get("Tools.LeastPrivilege.Status.Pending")
    };

    public string StatusColor => _evaluation.StatusColor;

    public bool HasMissingRequirements => _evaluation.HasMissingRequirements;

    public string MissingRequirementsDisplay => _evaluation.MissingRequirements.Count == 0
        ? Loc.Get("Common.None")
        : string.Join("; ", _evaluation.MissingRequirements);

    public string AllowedAuthenticationModesDisplay => _evaluation.Definition.AllowedAuthenticationModes.Count == 0
        ? Loc.Get("Tools.LeastPrivilege.CurrentAuthModes")
        : string.Join(", ", _evaluation.Definition.AllowedAuthenticationModes);

    public string RequiredCmdletsDisplay => FormatRequirementList(
        _evaluation.Definition.RequiredCmdletsAll,
        _evaluation.Definition.RequiredCmdletsAny);

    public string ExchangeRolesDisplay => FormatList(_evaluation.Definition.RecommendedExchangeRoles);

    public string GraphScopesDisplay => FormatList(_evaluation.Definition.RequiredGraphScopes);

    public string PurviewRolesDisplay => FormatList(_evaluation.Definition.RecommendedPurviewRoles);

    public string DefenderRolesDisplay => FormatList(_evaluation.Definition.RecommendedDefenderRoles);

    public string DependenciesDisplay => FormatList(_evaluation.Definition.Dependencies.Select(LocalizeDependency).ToArray());

    public string NotesDisplay => string.IsNullOrWhiteSpace(_evaluation.Definition.Notes)
        ? Loc.Get("Tools.LeastPrivilege.DefaultNotes")
        : GetFeatureText("Notes", _evaluation.Definition.Notes!);

    public string ValidationMessage => _evaluation.Status switch
    {
        LeastPrivilegeFeatureStatus.Available => Loc.Get("Tools.LeastPrivilege.Validation.Available"),
        LeastPrivilegeFeatureStatus.NeedsAdditionalSession => Loc.Get("Tools.LeastPrivilege.Validation.NeedsAdditionalSession"),
        LeastPrivilegeFeatureStatus.Blocked => Loc.Get("Tools.LeastPrivilege.Validation.Blocked"),
        _ => Loc.Get("Tools.LeastPrivilege.Validation.PendingSession")
    };

    private string GetFeatureText(string suffix, string fallback)
    {
        var key = $"Tools.LeastPrivilege.Feature.{_evaluation.FeatureId}.{suffix}";
        var value = Loc.Get(key);
        return string.Equals(value, key, StringComparison.Ordinal) ? fallback : value;
    }

    private static string FormatRequirementList(
        IReadOnlyList<string> requiredAll,
        IReadOnlyList<string> requiredAny)
    {
        if (requiredAll.Count == 0 && requiredAny.Count == 0)
        {
            return Loc.Get("Common.None");
        }

        var parts = new List<string>();
        if (requiredAll.Count > 0)
        {
            parts.Add(Loc.GetFormat("Tools.LeastPrivilege.RequiredAll", string.Join(", ", requiredAll)));
        }

        if (requiredAny.Count > 0)
        {
            parts.Add(Loc.GetFormat("Tools.LeastPrivilege.RequiredAny", string.Join(", ", requiredAny)));
        }

        return string.Join(" | ", parts);
    }

    private static string FormatList(IReadOnlyList<string> items)
        => items.Count == 0 ? Loc.Get("Common.None") : string.Join(", ", items);

    private static string LocalizeDependency(string dependency)
        => dependency switch
        {
            "Exchange Online session" => Loc.Get("Tools.LeastPrivilege.Dependency.ExchangeSession"),
            "Exchange Online Protection telemetry" => Loc.Get("Tools.LeastPrivilege.Dependency.EopTelemetry"),
            "Exchange transport configuration" => Loc.Get("Tools.LeastPrivilege.Dependency.TransportConfiguration"),
            "Purview / Security & Compliance PowerShell" => Loc.Get("Tools.LeastPrivilege.Dependency.PurviewPowerShell"),
            "Connect-IPPSSession during the initial connect flow" => Loc.Get("Tools.LeastPrivilege.Dependency.ConnectIPPSSession"),
            "Exchange Online Protection / Defender for Office 365" => Loc.Get("Tools.LeastPrivilege.Dependency.EopDefender"),
            "Microsoft Graph initial session" => Loc.Get("Tools.LeastPrivilege.Dependency.GraphInitialSession"),
            _ => dependency
        };
}

