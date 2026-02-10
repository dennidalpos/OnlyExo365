using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;

namespace ExchangeAdmin.Presentation.ViewModels;

public class ToolsViewModel : ViewModelBase
{
    private readonly IWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;

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

    public ToolsViewModel(IWorkerService workerService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;

        CheckPrerequisitesCommand = new AsyncRelayCommand(CheckPrerequisitesAsync, () => !IsChecking && !IsInstalling);
        InstallPowerShellCommand = new AsyncRelayCommand(InstallPowerShellAsync, () => !IsInstalling && !IsPowerShell7);
        InstallExchangeModuleCommand = new AsyncRelayCommand(() => InstallModuleAsync("ExchangeOnlineManagement"), () => !IsInstalling);
        InstallGraphModuleCommand = new AsyncRelayCommand(() => InstallModuleAsync("Microsoft.Graph.Authentication"), () => !IsInstalling);
        OpenPowerShellGitHubCommand = new RelayCommand(OpenPowerShellGitHub);
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

    public ICommand CheckPrerequisitesCommand { get; }
    public ICommand InstallPowerShellCommand { get; }
    public ICommand InstallExchangeModuleCommand { get; }
    public ICommand InstallGraphModuleCommand { get; }
    public ICommand OpenPowerShellGitHubCommand { get; }

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

    public async Task LoadAsync()
    {
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
        InstallStatus = "Verifica prerequisiti...";

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

                _shellViewModel.AddLog(LogLevel.Information,
                    $"[Prerequisites] PS={status.PowerShellVersion} (pwsh7={status.IsPowerShell7}), EXO={status.ExchangeModuleInstalled}, Graph={status.GraphModuleInstalled}");

                PrerequisiteSummary = $"PowerShell: {status.PowerShellVersion} | Exchange module: {(status.ExchangeModuleInstalled ? "OK" : "Missing")} | Graph module: {(status.GraphModuleInstalled ? "OK" : "Missing")}";

                var suggestions = new List<string>();
                if (!status.IsPowerShell7) suggestions.Add("Installare PowerShell 7 per compatibilità ottimale.");
                if (!status.ExchangeModuleInstalled) suggestions.Add("Installare/aggiornare ExchangeOnlineManagement.");
                if (!status.GraphModuleInstalled) suggestions.Add("Installare/aggiornare Microsoft.Graph.Authentication.");
                SuggestedActions = suggestions.Count == 0 ? "Prerequisiti completi. Nessuna azione richiesta." : string.Join(" ", suggestions);
            }
            else if (!result.WasCancelled)
            {
                ErrorMessage = result.Error?.Message ?? "Impossibile verificare i prerequisiti";
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
        InstallStatus = "Tentativo installazione PowerShell 7 tramite winget...";

        try
        {
            InstallProgress = 20;
            var result = await _workerService.InstallModuleAsync(
                new InstallModuleRequest { ModuleName = "PowerShell7_winget" },
                cancellationToken: cancellationToken);

            InstallProgress = 90;

            if (result.IsSuccess && result.Value != null)
            {
                if (result.Value.Success)
                {
                    _shellViewModel.AddLog(LogLevel.Information, "[ModuleInstall] PowerShell 7 installato tramite winget");
                    InstallStatus = "PowerShell 7 installato. Riavviare l'applicazione.";
                }
                else
                {
                    ErrorMessage = $"Installazione automatica fallita: {result.Value.Message}\n\n" +
                                   "Installare manualmente PowerShell 7:\n" +
                                   "1. Scaricare da https://github.com/PowerShell/PowerShell/releases\n" +
                                   "2. Oppure eseguire: winget install Microsoft.PowerShell";
                    _shellViewModel.AddLog(LogLevel.Warning, $"[ModuleInstall] PowerShell 7 install failed: {result.Value.Message}");
                }
            }
            else if (!result.WasCancelled)
            {
                ErrorMessage = result.Error?.Message ?? "Errore durante l'installazione";
            }

            InstallProgress = 100;
        }
        catch (Exception ex)
        {
            ErrorMessage = $"Errore: {ex.Message}\n\n" +
                           "Installare manualmente PowerShell 7:\n" +
                           "https://github.com/PowerShell/PowerShell/releases";
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
        InstallStatus = $"Installazione {moduleName}...";

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
                    _shellViewModel.AddLog(LogLevel.Information, $"[ModuleInstall] {moduleName} installato con successo");
                    InstallStatus = $"{moduleName} installato.";
                    await CheckPrerequisitesAsync(CancellationToken.None);
                }
                else
                {
                    ErrorMessage = $"Installazione {moduleName} fallita: {result.Value.Message}\n\n" +
                                   $"Provare manualmente in PowerShell:\n" +
                                   $"Install-Module {moduleName} -Force -AllowClobber -Scope CurrentUser";
                    _shellViewModel.AddLog(LogLevel.Error, $"[ModuleInstall] {moduleName} install failed: {result.Value.Message}");
                }
            }
            else if (!result.WasCancelled)
            {
                ErrorMessage = result.Error?.Message ?? $"Errore installazione {moduleName}";
            }

            InstallProgress = 100;
        }
        catch (Exception ex)
        {
            ErrorMessage = $"Errore: {ex.Message}\n\n" +
                           $"Installare manualmente:\nInstall-Module {moduleName} -Force -AllowClobber -Scope CurrentUser";
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
            _shellViewModel.AddLog(LogLevel.Error, $"Impossibile aprire il browser: {ex.Message}");
        }
    }
}
