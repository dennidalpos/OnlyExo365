using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;

namespace ExchangeAdmin.Presentation.ViewModels;

public sealed class MailboxRestoreViewModel : ViewModelBase
{
    private readonly IMailboxesWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;

    private string? _restoreSourceIdentity;
    private string? _restoreTargetMailbox;
    private bool _restoreAllowLegacyDnMismatch;
    private bool _isRestoringMailbox;
    private RestoreMailboxResponse? _restoreMailboxResponse;
    private string? _restoreMailboxErrorMessage;

    public MailboxRestoreViewModel(IMailboxesWorkerService workerService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;
        RestoreMailboxCommand = new AsyncRelayCommand(RestoreMailboxAsync, () => !IsRestoringMailbox);
    }

    public string? RestoreSourceIdentity
    {
        get => _restoreSourceIdentity;
        set => SetProperty(ref _restoreSourceIdentity, value);
    }

    public string? RestoreTargetMailbox
    {
        get => _restoreTargetMailbox;
        set => SetProperty(ref _restoreTargetMailbox, value);
    }

    public bool RestoreAllowLegacyDnMismatch
    {
        get => _restoreAllowLegacyDnMismatch;
        set => SetProperty(ref _restoreAllowLegacyDnMismatch, value);
    }

    public bool IsRestoringMailbox
    {
        get => _isRestoringMailbox;
        private set
        {
            if (SetProperty(ref _isRestoringMailbox, value))
            {
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public RestoreMailboxResponse? RestoreMailboxResponse
    {
        get => _restoreMailboxResponse;
        private set
        {
            if (SetProperty(ref _restoreMailboxResponse, value))
            {
                _restoreMailboxErrorMessage = null;
                OnPropertyChanged(nameof(HasRestoreMailboxResponse));
                OnPropertyChanged(nameof(RestoreScenarioText));
                OnPropertyChanged(nameof(RestoreActionText));
                OnPropertyChanged(nameof(RestoreStatusText));
                OnPropertyChanged(nameof(RestoreStatusDetail));
                OnPropertyChanged(nameof(RestorePercentComplete));
                OnPropertyChanged(nameof(RestoreRequestGuid));
                OnPropertyChanged(nameof(RestoreErrorCodeText));
                OnPropertyChanged(nameof(RestoreErrorMessage));
                OnPropertyChanged(nameof(HasRestoreError));
                OnPropertyChanged(nameof(RestoreProgressValue));
                OnPropertyChanged(nameof(IsRestoreProgressIndeterminate));
                OnPropertyChanged(nameof(RestoreProgressText));
            }
        }
    }

    public bool HasRestoreMailboxResponse => RestoreMailboxResponse != null;

    public string RestoreScenarioText => RestoreMailboxResponse == null
        ? "-"
        : RestoreMailboxResponse.Scenario switch
        {
            RestoreMailboxScenario.SoftDeleted => "Soft-deleted",
            RestoreMailboxScenario.Inactive => "Inactive",
            RestoreMailboxScenario.HardDeleted => "Hard-deleted",
            RestoreMailboxScenario.Existing => "Existing mailbox",
            RestoreMailboxScenario.NotFound => "Not found",
            _ => "Unknown"
        };

    public string RestoreActionText => RestoreMailboxResponse?.Action ?? "-";

    public string RestoreStatusText => RestoreMailboxResponse == null
        ? "-"
        : RestoreMailboxResponse.Status switch
        {
            RestoreMailboxStatus.InProgress => "In progress",
            RestoreMailboxStatus.Completed => "Completed",
            RestoreMailboxStatus.Failed => "Error",
            _ => "Not started"
        };

    public string RestoreStatusDetail => RestoreMailboxResponse?.StatusDetail ?? "-";
    public int? RestorePercentComplete => RestoreMailboxResponse?.PercentComplete;
    public string RestoreRequestGuid => RestoreMailboxResponse?.RequestGuid ?? "-";
    public string RestoreErrorCodeText => RestoreMailboxResponse?.Error?.Code.ToString() ?? "-";

    public string? RestoreErrorMessage
    {
        get => RestoreMailboxResponse?.Error?.Message ?? _restoreMailboxErrorMessage;
        private set
        {
            if (SetProperty(ref _restoreMailboxErrorMessage, value))
            {
                OnPropertyChanged(nameof(HasRestoreError));
            }
        }
    }

    public bool HasRestoreError => !string.IsNullOrWhiteSpace(RestoreErrorMessage);
    public int RestoreProgressValue => RestoreMailboxResponse?.PercentComplete ?? 0;

    public bool IsRestoreProgressIndeterminate =>
        RestoreMailboxResponse?.Status == RestoreMailboxStatus.InProgress &&
        (RestoreMailboxResponse?.PercentComplete.HasValue == false);

    public string RestoreProgressText => RestoreMailboxResponse?.PercentComplete.HasValue == true
        ? $"{RestoreMailboxResponse.PercentComplete}% completed"
        : "Progress not available";

    public ICommand RestoreMailboxCommand { get; }

    public void Reset(string? identity)
    {
        RestoreSourceIdentity = identity;
        RestoreTargetMailbox = null;
        RestoreAllowLegacyDnMismatch = false;
        RestoreMailboxResponse = null;
        RestoreErrorMessage = null;
    }

    private async Task RestoreMailboxAsync(CancellationToken cancellationToken)
    {
        var sourceIdentity = RestoreSourceIdentity?.Trim();
        var targetMailbox = RestoreTargetMailbox?.Trim();

        if (string.IsNullOrWhiteSpace(sourceIdentity))
        {
            RestoreMailboxResponse = null;
            RestoreErrorMessage = "Specify the source mailbox UPN or GUID.";
            return;
        }

        IsRestoringMailbox = true;
        RestoreMailboxResponse = null;
        RestoreErrorMessage = null;

        try
        {
            var request = new RestoreMailboxRequest
            {
                SourceIdentity = sourceIdentity,
                TargetMailbox = string.IsNullOrWhiteSpace(targetMailbox) ? null : targetMailbox,
                AllowLegacyDnMismatch = RestoreAllowLegacyDnMismatch
            };

            _shellViewModel.AddLog(LogLevel.Information, $"Starting mailbox restore for {request.SourceIdentity}...");

            var result = await _workerService.RestoreMailboxAsync(request, cancellationToken: cancellationToken);
            if (result.IsSuccess && result.Value != null)
            {
                RestoreMailboxResponse = result.Value;
                if (RestoreMailboxResponse.Error != null)
                {
                    _shellViewModel.AddLog(LogLevel.Warning, $"Mailbox restore completed with an error: {RestoreMailboxResponse.Error.Message}");
                }
                else
                {
                    _shellViewModel.AddLog(LogLevel.Information, $"Mailbox restore started: {RestoreMailboxResponse.Status}");
                }
            }
            else
            {
                RestoreErrorMessage = result.IsSuccess && result.Value == null
                    ? "Restore response is not available."
                    : result.Error?.Message ?? "Unable to start the mailbox restore.";

                _shellViewModel.AddLog(LogLevel.Error, $"Mailbox restore failed: {RestoreErrorMessage}");
            }
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            RestoreErrorMessage = ex.Message;
            _shellViewModel.AddLog(LogLevel.Error, $"Mailbox restore error: {ex.Message}");
        }
        finally
        {
            IsRestoringMailbox = false;
        }
    }
}
