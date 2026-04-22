using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Shell.Helpers;

namespace OnlyExo365.Shell.ViewModels;

public sealed class DistributionListSettingsEditorViewModel : ViewModelBase
{
    private readonly IDistributionListsWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;
    private readonly Func<DistributionListDetailsDto?> _getSelectedDetails;
    private readonly Action<string?> _setErrorMessage;

    private bool _isSavingSettings;
    private bool _isInitializingSettings;
    private bool _allowExternalSenders;
    private bool _originalAllowExternalSenders;
    private bool _hasPendingSettingsChanges;
    private string? _newAcceptedSender;
    private string? _newRejectedSender;
    private List<string> _originalAcceptMessagesOnlyFrom = new();
    private List<string> _originalRejectMessagesFrom = new();

    public DistributionListSettingsEditorViewModel(
        IDistributionListsWorkerService workerService,
        ShellViewModel shellViewModel,
        Func<DistributionListDetailsDto?> getSelectedDetails,
        Action<string?> setErrorMessage)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;
        _getSelectedDetails = getSelectedDetails;
        _setErrorMessage = setErrorMessage;

        SaveSettingsCommand = new AsyncRelayCommand(SaveSettingsAsync, () => HasPendingSettingsChanges && !IsSavingSettings);
        DiscardSettingsCommand = new RelayCommand(DiscardSettingsChanges, () => HasPendingSettingsChanges);
        AddAcceptSenderCommand = new RelayCommand(AddAcceptSender, () => CanAddAcceptSender);
        RemoveAcceptSenderCommand = new RelayCommand<string>(RemoveAcceptSender, sender => CanRemoveAcceptSender(sender));
        AddRejectSenderCommand = new RelayCommand(AddRejectSender, () => CanAddRejectSender);
        RemoveRejectSenderCommand = new RelayCommand<string>(RemoveRejectSender, sender => CanRemoveRejectSender(sender));
    }

    public ObservableCollection<string> AcceptMessagesOnlyFrom { get; } = new();
    public ObservableCollection<string> RejectMessagesFrom { get; } = new();

    public bool IsSavingSettings
    {
        get => _isSavingSettings;
        private set
        {
            if (SetProperty(ref _isSavingSettings, value))
            {
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool AllowExternalSenders
    {
        get => _allowExternalSenders;
        set
        {
            if (SetProperty(ref _allowExternalSenders, value))
            {
                UpdatePendingSettingsChanges();
            }
        }
    }

    public bool HasPendingSettingsChanges
    {
        get => _hasPendingSettingsChanges;
        private set
        {
            if (SetProperty(ref _hasPendingSettingsChanges, value))
            {
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public string? NewAcceptedSender
    {
        get => _newAcceptedSender;
        set
        {
            if (SetProperty(ref _newAcceptedSender, value))
            {
                OnPropertyChanged(nameof(CanAddAcceptSender));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public string? NewRejectedSender
    {
        get => _newRejectedSender;
        set
        {
            if (SetProperty(ref _newRejectedSender, value))
            {
                OnPropertyChanged(nameof(CanAddRejectSender));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool CanEditSettings => HasDetails && !IsMicrosoft365Group && (IsDynamicGroup
        ? _shellViewModel.IsFeatureAvailable(f => f.CanSetDynamicDistributionGroup)
        : _shellViewModel.IsFeatureAvailable(f => f.CanSetDistributionGroup));

    public bool CanEditExternalSenders => CanEditSettings && (IsDynamicGroup
        ? _shellViewModel.IsFeatureAvailable(f => f.CanSetDynamicDistributionGroupRequireSenderAuthentication)
        : _shellViewModel.IsFeatureAvailable(f => f.CanSetDistributionGroupRequireSenderAuthentication));

    public bool CanEditAcceptMessagesOnlyFrom => CanEditSettings && (IsDynamicGroup
        ? _shellViewModel.IsFeatureAvailable(f => f.CanSetDynamicDistributionGroupAcceptMessagesOnlyFrom)
        : _shellViewModel.IsFeatureAvailable(f => f.CanSetDistributionGroupAcceptMessagesOnlyFrom));

    public bool CanEditRejectMessagesFrom => CanEditSettings && (IsDynamicGroup
        ? _shellViewModel.IsFeatureAvailable(f => f.CanSetDynamicDistributionGroupRejectMessagesFrom)
        : _shellViewModel.IsFeatureAvailable(f => f.CanSetDistributionGroupRejectMessagesFrom));

    public bool CanAddAcceptSender => CanEditAcceptMessagesOnlyFrom && !string.IsNullOrWhiteSpace(NewAcceptedSender);
    public bool CanAddRejectSender => CanEditRejectMessagesFrom && !string.IsNullOrWhiteSpace(NewRejectedSender);

    public ICommand SaveSettingsCommand { get; }
    public ICommand DiscardSettingsCommand { get; }
    public ICommand AddAcceptSenderCommand { get; }
    public ICommand RemoveAcceptSenderCommand { get; }
    public ICommand AddRejectSenderCommand { get; }
    public ICommand RemoveRejectSenderCommand { get; }

    private bool HasDetails => _getSelectedDetails() != null;
    private bool IsDynamicGroup => string.Equals(_getSelectedDetails()?.GroupType, "Dynamic", StringComparison.OrdinalIgnoreCase);
    private bool IsMicrosoft365Group => string.Equals(_getSelectedDetails()?.GroupType, "Microsoft365", StringComparison.OrdinalIgnoreCase);

    public void InitializeFromDetails(DistributionListDetailsDto? details)
    {
        _isInitializingSettings = true;

        if (details != null)
        {
            AllowExternalSenders = !details.RequireSenderAuthenticationEnabled;
            _originalAllowExternalSenders = AllowExternalSenders;
            _originalAcceptMessagesOnlyFrom = DistributionListViewModelSupport.NormalizeSenderList(details.AcceptMessagesOnlyFrom).ToList();
            _originalRejectMessagesFrom = DistributionListViewModelSupport.NormalizeSenderList(details.RejectMessagesFrom).ToList();
            DistributionListViewModelSupport.ResetObservableList(AcceptMessagesOnlyFrom, _originalAcceptMessagesOnlyFrom);
            DistributionListViewModelSupport.ResetObservableList(RejectMessagesFrom, _originalRejectMessagesFrom);
        }
        else
        {
            AllowExternalSenders = false;
            _originalAllowExternalSenders = false;
            _originalAcceptMessagesOnlyFrom = new List<string>();
            _originalRejectMessagesFrom = new List<string>();
            AcceptMessagesOnlyFrom.Clear();
            RejectMessagesFrom.Clear();
        }

        NewAcceptedSender = string.Empty;
        NewRejectedSender = string.Empty;
        _isInitializingSettings = false;
        HasPendingSettingsChanges = false;
        RaiseCapabilityPropertiesChanged();
    }

    public void HandleShellPropertyChanged()
    {
        RaiseCapabilityPropertiesChanged();
        UpdatePendingSettingsChanges();
        CommandManager.InvalidateRequerySuggested();
    }

    private async Task SaveSettingsAsync(CancellationToken cancellationToken)
    {
        var selectedDetails = _getSelectedDetails();
        if (selectedDetails == null || !CanEditSettings)
        {
            return;
        }

        if (!ConfirmMutation(
                "Updating settings group",
                selectedDetails.Identity,
                "Update allowed/blocked senders and sender authentication for the distribution list.",
                "Confirm group update"))
        {
            return;
        }

        IsSavingSettings = true;
        _setErrorMessage(null);

        try
        {
            var request = new SetDistributionListSettingsRequest
            {
                Identity = selectedDetails.Identity,
                GroupType = DistributionListViewModelSupport.MapGroupTypeForWorker(selectedDetails.GroupType),
                RequireSenderAuthenticationEnabled = CanEditExternalSenders ? !AllowExternalSenders : null,
                AcceptMessagesOnlyFrom = CanEditAcceptMessagesOnlyFrom ? DistributionListViewModelSupport.NormalizeSenderList(AcceptMessagesOnlyFrom).ToList() : null,
                RejectMessagesFrom = CanEditRejectMessagesFrom ? DistributionListViewModelSupport.NormalizeSenderList(RejectMessagesFrom).ToList() : null
            };

            var result = await _workerService.SetDistributionListSettingsAsync(request, cancellationToken: cancellationToken);
            if (result.IsSuccess)
            {
                if (CanEditExternalSenders)
                {
                    selectedDetails.RequireSenderAuthenticationEnabled = !AllowExternalSenders;
                    _originalAllowExternalSenders = AllowExternalSenders;
                }

                if (CanEditAcceptMessagesOnlyFrom)
                {
                    selectedDetails.AcceptMessagesOnlyFrom = DistributionListViewModelSupport.NormalizeSenderList(AcceptMessagesOnlyFrom).ToList();
                    _originalAcceptMessagesOnlyFrom = selectedDetails.AcceptMessagesOnlyFrom.ToList();
                }

                if (CanEditRejectMessagesFrom)
                {
                    selectedDetails.RejectMessagesFrom = DistributionListViewModelSupport.NormalizeSenderList(RejectMessagesFrom).ToList();
                    _originalRejectMessagesFrom = selectedDetails.RejectMessagesFrom.ToList();
                }

                HasPendingSettingsChanges = false;
                return;
            }

            if (!result.WasCancelled)
            {
                var errorMessage = result.Error?.Message ?? "Unable to update group settings.";
                _setErrorMessage(errorMessage);
                _shellViewModel.AddLog(LogLevel.Error, $"Save group settings failed: {errorMessage}");
            }
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            _setErrorMessage(ex.Message);
            _shellViewModel.AddLog(LogLevel.Error, $"Save group settings error: {ex.Message}");
        }
        finally
        {
            IsSavingSettings = false;
        }
    }

    private void DiscardSettingsChanges()
    {
        if (_isInitializingSettings)
        {
            return;
        }

        AllowExternalSenders = _originalAllowExternalSenders;
        DistributionListViewModelSupport.ResetObservableList(AcceptMessagesOnlyFrom, _originalAcceptMessagesOnlyFrom);
        DistributionListViewModelSupport.ResetObservableList(RejectMessagesFrom, _originalRejectMessagesFrom);
        HasPendingSettingsChanges = false;
    }

    private void UpdatePendingSettingsChanges()
    {
        if (_isInitializingSettings)
        {
            return;
        }

        var externalChanged = CanEditExternalSenders && AllowExternalSenders != _originalAllowExternalSenders;
        var acceptChanged = CanEditAcceptMessagesOnlyFrom && !DistributionListViewModelSupport.SenderListEquals(AcceptMessagesOnlyFrom, _originalAcceptMessagesOnlyFrom);
        var rejectChanged = CanEditRejectMessagesFrom && !DistributionListViewModelSupport.SenderListEquals(RejectMessagesFrom, _originalRejectMessagesFrom);
        HasPendingSettingsChanges = externalChanged || acceptChanged || rejectChanged;
    }

    private void AddAcceptSender()
    {
        if (!CanAddAcceptSender)
        {
            return;
        }

        var normalized = NewAcceptedSender!.Trim();
        if (!AcceptMessagesOnlyFrom.Any(sender => string.Equals(sender, normalized, StringComparison.OrdinalIgnoreCase)))
        {
            AcceptMessagesOnlyFrom.Add(normalized);
            UpdatePendingSettingsChanges();
        }

        NewAcceptedSender = string.Empty;
    }

    private void RemoveAcceptSender(string? sender)
    {
        if (!CanRemoveAcceptSender(sender))
        {
            return;
        }

        AcceptMessagesOnlyFrom.Remove(sender!);
        UpdatePendingSettingsChanges();
    }

    private void AddRejectSender()
    {
        if (!CanAddRejectSender)
        {
            return;
        }

        var normalized = NewRejectedSender!.Trim();
        if (!RejectMessagesFrom.Any(sender => string.Equals(sender, normalized, StringComparison.OrdinalIgnoreCase)))
        {
            RejectMessagesFrom.Add(normalized);
            UpdatePendingSettingsChanges();
        }

        NewRejectedSender = string.Empty;
    }

    private void RemoveRejectSender(string? sender)
    {
        if (!CanRemoveRejectSender(sender))
        {
            return;
        }

        RejectMessagesFrom.Remove(sender!);
        UpdatePendingSettingsChanges();
    }

    private bool CanRemoveAcceptSender(string? sender) => CanEditAcceptMessagesOnlyFrom && !string.IsNullOrWhiteSpace(sender);
    private bool CanRemoveRejectSender(string? sender) => CanEditRejectMessagesFrom && !string.IsNullOrWhiteSpace(sender);

    private void RaiseCapabilityPropertiesChanged()
    {
        OnPropertyChanged(nameof(CanEditSettings));
        OnPropertyChanged(nameof(CanEditExternalSenders));
        OnPropertyChanged(nameof(CanEditAcceptMessagesOnlyFrom));
        OnPropertyChanged(nameof(CanEditRejectMessagesFrom));
        OnPropertyChanged(nameof(CanAddAcceptSender));
        OnPropertyChanged(nameof(CanAddRejectSender));
    }
}

