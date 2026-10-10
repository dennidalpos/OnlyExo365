using OnlyExo365.Shell.Localization;
using System.ComponentModel;
using System.Windows.Input;
using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Shell.Helpers;

namespace OnlyExo365.Shell.ViewModels;

public class ResourcesPageViewModel : ViewModelBase
{
    private readonly IResourcesWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;

    private bool _isSaving;
    private int _detailsLoadingCount;
    private ResourceMailboxDetailsDto? _loadedDetails;
    private string? _errorMessage;

    public ResourcesPageViewModel(IResourcesWorkerService workerService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;

        Editor = new ResourceMailboxEditorViewModel();
        List = new ResourcesListStateViewModel(workerService, shellViewModel, OnResourceSelected, SetErrorMessage);

        List.CanReplaceDraft = CanReplaceDraft;

        Editor.PropertyChanged += OnEditorPropertyChanged;
        List.PropertyChanged += OnListPropertyChanged;
        _shellViewModel.PropertyChanged += OnShellPropertyChanged;

        SaveCommand = new AsyncRelayCommand(SaveAsync, () => CanSave);
        NewResourceCommand = new RelayCommand(BeginCreate, () => !IsSaving);
    }

    public ResourcesListStateViewModel List { get; }

    public ResourceMailboxEditorViewModel Editor { get; }

    public bool IsLoading => List.IsLoading;

    public bool IsEditorBusy => IsLoading || IsSaving || _detailsLoadingCount > 0;

    public bool IsSaving
    {
        get => _isSaving;
        private set
        {
            if (SetProperty(ref _isSaving, value))
            {
                OnPropertyChanged(nameof(CanSave));
                OnPropertyChanged(nameof(IsEditorBusy));
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

    private bool CanReplaceDraft()
    {
        if (IsSaving || !HasPendingChanges) return true;
        if (!ErrorDialogService.ConfirmDiscardChanges()) return false;
        DiscardDraft();
        return true;
    }

    public void DiscardDraft()
    {
        if (_loadedDetails != null) Editor.ApplyDetails(_loadedDetails);
        else Editor.BeginCreate();
    }

    private void BeginCreate()
    {
        if (!CanReplaceDraft()) return;
        _loadedDetails = null;
        Editor.BeginCreate();
        List.ClearSelection();
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

        _detailsLoadingCount++;
        OnPropertyChanged(nameof(IsEditorBusy));
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
                SetErrorMessage(result.Error?.Message ?? Loc.Get("Presentation.UnableToLoadResourceDetails"));
                return;
            }

            if (ReferenceEquals(List.SelectedResource, selected))
            {
                _loadedDetails = result.Value;
                Editor.ApplyDetails(result.Value);
            }
        }
        catch (Exception ex)
        {
            SetErrorMessage(ex.Message);
        }
        finally
        {
            _detailsLoadingCount--;
            OnPropertyChanged(nameof(IsEditorBusy));
        }
    }

    private async Task SaveAsync(CancellationToken cancellationToken)
    {
        if (!CanSave)
        {
            return;
        }

        var request = Editor.BuildUpsertRequest();
        var primarySmtpAddress = request.PrimarySmtpAddress;
        var permissionActions = Editor.BuildPermissionDelta();
        if (!ConfirmMutation(
                string.IsNullOrWhiteSpace(Editor.ResourceIdentity) ? Loc.Get("Presentation.CreatingResourceMailbox") : Loc.Get("Presentation.UpdatingResourceMailbox"),
                primarySmtpAddress,
                permissionActions.Count > 0
                    ? Loc.GetFormat("Resources.DelegationImpact", permissionActions.Count)
                    : Loc.Get("Presentation.SaveResourceMailboxAttributesAndConfiguration"),
                Loc.Get("Presentation.ConfirmResourceMailboxSave")))
        {
            return;
        }

        IsSaving = true;
        SetErrorMessage(null);

        try
        {
            var saveResult = await _workerService.UpsertResourceMailboxAsync(
                request,
                cancellationToken: cancellationToken);

            if (!saveResult.IsSuccess || saveResult.Value == null)
            {
                SetErrorMessage(saveResult.Error?.Message ?? Loc.Get("Presentation.UnableToSaveTheResourceMailbox"));
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
                    SetErrorMessage(permissionResult.Error?.Message ?? Loc.Get("Presentation.MailboxSavedButDelegationUpdateFailed"));
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
        OnPropertyChanged(nameof(IsEditorBusy));
        CommandManager.InvalidateRequerySuggested();
    }

    private void OnListPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(ResourcesListStateViewModel.IsLoading))
        {
            OnPropertyChanged(nameof(IsLoading));
        }

        OnPropertyChanged(nameof(CanSave));
        OnPropertyChanged(nameof(IsEditorBusy));
        CommandManager.InvalidateRequerySuggested();
    }

    private void OnShellPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName is not (nameof(ShellViewModel.ExchangeState) or nameof(ShellViewModel.IsExchangeConnected)))
        {
            return;
        }

        OnPropertyChanged(nameof(CanSave));
        OnPropertyChanged(nameof(IsEditorBusy));
        CommandManager.InvalidateRequerySuggested();
    }
    public override void ReleaseUiResources()
    {
        base.ReleaseUiResources();
        List.ReleaseUiResources();
        Editor.ReleaseUiResources();
    }
}


