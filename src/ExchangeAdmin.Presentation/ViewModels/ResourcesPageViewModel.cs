using System.ComponentModel;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;

namespace ExchangeAdmin.Presentation.ViewModels;

public class ResourcesPageViewModel : ViewModelBase
{
    private readonly IResourcesWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;

    private bool _isSaving;
    private string? _errorMessage;

    public ResourcesPageViewModel(IResourcesWorkerService workerService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;

        Editor = new ResourceMailboxEditorViewModel();
        List = new ResourcesListStateViewModel(workerService, shellViewModel, OnResourceSelected, SetErrorMessage);

        Editor.PropertyChanged += OnEditorPropertyChanged;
        List.PropertyChanged += OnListPropertyChanged;
        _shellViewModel.PropertyChanged += OnShellPropertyChanged;

        SaveCommand = new AsyncRelayCommand(SaveAsync, () => CanSave);
        NewResourceCommand = new RelayCommand(BeginCreate, () => !IsSaving);
    }

    public ResourcesListStateViewModel List { get; }

    public ResourceMailboxEditorViewModel Editor { get; }

    public bool IsLoading => List.IsLoading;

    public bool IsSaving
    {
        get => _isSaving;
        private set
        {
            if (SetProperty(ref _isSaving, value))
            {
                OnPropertyChanged(nameof(CanSave));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public bool HasPendingChanges => Editor.HasPendingChanges;

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

    public bool CanSave => !IsLoading && !IsSaving && _shellViewModel.IsExchangeConnected && Editor.CanSave;

    public ICommand RefreshCommand => List.RefreshCommand;
    public ICommand LoadMoreCommand => List.LoadMoreCommand;
    public ICommand SaveCommand { get; }
    public ICommand NewResourceCommand { get; }

    public async Task LoadAsync()
    {
        await List.LoadAsync(CancellationToken.None);
    }

    private void BeginCreate()
    {
        List.ClearSelection();
        Editor.BeginCreate();
        SetErrorMessage(null);
    }

    private async void OnResourceSelected(ResourceMailboxListItemDto? selected)
    {
        if (selected == null)
        {
            return;
        }

        await LoadDetailsAsync(selected);
    }

    private async Task LoadDetailsAsync(ResourceMailboxListItemDto selected)
    {
        SetErrorMessage(null);

        try
        {
            var result = await _workerService.GetResourceMailboxDetailsAsync(
                new GetResourceMailboxDetailsRequest
                {
                    Identity = selected.Identity
                },
                cancellationToken: CancellationToken.None);

            if (!result.IsSuccess || result.Value == null)
            {
                SetErrorMessage(result.Error?.Message ?? "Unable to load resource details.");
                return;
            }

            Editor.ApplyDetails(result.Value);
        }
        catch (Exception ex)
        {
            SetErrorMessage(ex.Message);
        }
    }

    private async Task SaveAsync(CancellationToken cancellationToken)
    {
        if (!CanSave)
        {
            return;
        }

        var primarySmtpAddress = Editor.PrimarySmtpAddress.Trim();
        var permissionActions = Editor.BuildPermissionDelta();
        if (!ConfirmMutation(
                string.IsNullOrWhiteSpace(Editor.ResourceIdentity) ? "Creating resource mailbox" : "Updating resource mailbox",
                primarySmtpAddress,
                permissionActions.Count > 0
                    ? $"Save the resource mailbox and apply {permissionActions.Count} delegation changes."
                    : "Save resource mailbox attributes and configuration.",
                "Confirm resource mailbox save"))
        {
            return;
        }

        IsSaving = true;
        SetErrorMessage(null);

        try
        {
            var saveResult = await _workerService.UpsertResourceMailboxAsync(
                Editor.BuildUpsertRequest(),
                cancellationToken: cancellationToken);

            if (!saveResult.IsSuccess || saveResult.Value == null)
            {
                SetErrorMessage(saveResult.Error?.Message ?? "Unable to save the resource mailbox.");
                return;
            }

            if (permissionActions.Count > 0)
            {
                var permissionResult = await _workerService.ApplyPermissionsDeltaPlanAsync(
                    new ApplyPermissionsDeltaPlanRequest
                    {
                        Identity = saveResult.Value.Identity,
                        Actions = permissionActions
                    },
                    cancellationToken: cancellationToken);

                if (!permissionResult.IsSuccess)
                {
                    SetErrorMessage(permissionResult.Error?.Message ?? "Mailbox saved but delegation update failed");
                    return;
                }
            }

            _shellViewModel.AddLog(LogLevel.Information, $"Resource mailbox saved: {Editor.DisplayName} ({Editor.ResourceType})", "Resources");
            Editor.AcceptCurrentStateAsOriginal(saveResult.Value.Identity);
            await List.RefreshAsync(cancellationToken);
            await TryReselectSavedResourceAsync(primarySmtpAddress);
        }
        catch (Exception ex)
        {
            SetErrorMessage(ex.Message);
        }
        finally
        {
            IsSaving = false;
        }
    }

    private async Task TryReselectSavedResourceAsync(string primarySmtpAddress)
    {
        var match = List.Resources.FirstOrDefault(resource =>
            string.Equals(resource.PrimarySmtpAddress, primarySmtpAddress, StringComparison.OrdinalIgnoreCase));

        if (match == null)
        {
            return;
        }

        List.SelectedResource = match;
        await LoadDetailsAsync(match);
    }

    private void SetErrorMessage(string? errorMessage)
    {
        ErrorMessage = errorMessage;
    }

    private void OnEditorPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(ResourceMailboxEditorViewModel.HasPendingChanges))
        {
            OnPropertyChanged(nameof(HasPendingChanges));
        }

        OnPropertyChanged(nameof(CanSave));
        CommandManager.InvalidateRequerySuggested();
    }

    private void OnListPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(ResourcesListStateViewModel.IsLoading))
        {
            OnPropertyChanged(nameof(IsLoading));
        }

        OnPropertyChanged(nameof(CanSave));
        CommandManager.InvalidateRequerySuggested();
    }

    private void OnShellPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName is not (nameof(ShellViewModel.ExchangeState) or nameof(ShellViewModel.IsExchangeConnected)))
        {
            return;
        }

        OnPropertyChanged(nameof(CanSave));
        CommandManager.InvalidateRequerySuggested();
    }
}

