using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;

namespace ExchangeAdmin.Presentation.ViewModels;

public partial class MobileDevicesViewModel
{
    private async Task RefreshAsync(CancellationToken cancellationToken)
    {
        var hasWorkspaceData = HasWorkspaceData;

        if (!_shellViewModel.IsExchangeConnected)
        {
            ResetDisconnectedState();
            return;
        }

        if (!await EnsureCapabilityStateAsync(cancellationToken))
        {
            return;
        }

        IsLoading = true;
        ErrorMessage = null;
        var refreshPageSize = GetRefreshPageSize(Devices.Count);
        _currentSkip = 0;
        ClearLoadingProgress();

        try
        {
            var result = await _workerService.GetMobileDevicesAsync(
                BuildRequest(0, refreshPageSize),
                eventHandler: HandleWorkerEvent,
                cancellationToken: cancellationToken);

            if (!result.IsSuccess || result.Value == null)
            {
                var errorMessage = result.Error?.Message ?? "Unable to load mobile devices";
                if (hasWorkspaceData)
                {
                    ErrorMessage = errorMessage;
                }
                else
                {
                    _shellViewModel.ShowPageLoadFailedAlert(AlertPage, errorMessage);
                }
                return;
            }

            Devices.ReplaceAll(result.Value.Devices);
            TotalCount = result.Value.TotalCount;
            HasMore = result.Value.HasMore;
            IsTotalCountExact = result.Value.IsTotalCountExact;
            _currentSkip = Devices.Count;
            TryRestoreSelection();
            _shellViewModel.ClearPageAlert(AlertPage);
        }
        catch (Exception ex)
        {
            if (hasWorkspaceData)
            {
                ErrorMessage = ex.Message;
            }
            else
            {
                _shellViewModel.ShowPageLoadFailedAlert(AlertPage, ex.Message);
            }
        }
        finally
        {
            IsLoading = false;
        }
    }

    private async Task LoadMoreAsync(CancellationToken cancellationToken)
    {
        if (!HasMore || !IsMobileDevicesFeatureAvailable)
        {
            return;
        }

        IsLoading = true;
        ErrorMessage = null;
        ClearLoadingProgress();

        try
        {
            var result = await _workerService.GetMobileDevicesAsync(
                BuildRequest(_currentSkip),
                eventHandler: HandleWorkerEvent,
                cancellationToken: cancellationToken);

            if (!result.IsSuccess || result.Value == null)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to load mobile devices";
                return;
            }

            foreach (var item in result.Value.Devices)
            {
                Devices.Add(item);
            }

            TotalCount = result.Value.TotalCount;
            HasMore = result.Value.HasMore;
            IsTotalCountExact = result.Value.IsTotalCountExact;
            _currentSkip = Devices.Count;
            OnPropertyChanged(nameof(StatusText));
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
        finally
        {
            IsLoading = false;
        }
    }

    private async Task LoadSelectedDeviceAsync(MobileDeviceListItemDto selected, CancellationToken cancellationToken)
    {
        if (!_shellViewModel.IsExchangeConnected || !IsMobileDevicesFeatureAvailable)
        {
            return;
        }

        IsLoadingSelection = true;
        ErrorMessage = null;
        ClearLoadingProgress();

        try
        {
            var requestedIdentity = selected.Identity;
            var detailsResult = await _workerService.GetMobileDeviceDetailsAsync(
                new GetMobileDeviceDetailsRequest
                {
                    Identity = selected.Identity,
                    MailboxIdentity = selected.MailboxIdentity
                },
                eventHandler: HandleWorkerEvent,
                cancellationToken: cancellationToken);

            if (!detailsResult.IsSuccess || detailsResult.Value?.Device == null)
            {
                if (IsStillSelected(requestedIdentity))
                {
                    ErrorMessage = detailsResult.Error?.Message ?? "Unable to load mobile device details.";
                }

                return;
            }

            if (!IsStillSelected(requestedIdentity))
            {
                return;
            }

            ApplySelectedDeviceDetails(detailsResult.Value.Device);

            if (Policies.Count == 0 && CanLoadPolicies)
            {
                await LoadPoliciesAsync(cancellationToken);

                if (IsStillSelected(requestedIdentity))
                {
                    SelectedMailboxPolicyIdentity = ResolveSelectedPolicy(SelectedDevice);
                }
            }
        }
        catch (Exception ex)
        {
            if (IsStillSelected(selected.Identity))
            {
                ErrorMessage = ex.Message;
            }
        }
        finally
        {
            ClearLoadingProgress();
            IsLoadingSelection = false;
        }
    }

    private async Task LoadPoliciesAsync(CancellationToken cancellationToken)
    {
        if (!CanLoadPolicies)
        {
            Policies.Clear();
            SelectedMailboxPolicyIdentity = null;
            return;
        }

        var result = await _workerService.GetMobileDeviceMailboxPoliciesAsync(
            eventHandler: HandleWorkerEvent,
            cancellationToken: cancellationToken);

        if (!result.IsSuccess || result.Value == null)
        {
            ErrorMessage = result.Error?.Message ?? "Unable to load ActiveSync mailbox policies.";
            Policies.Clear();
            return;
        }

        Policies.ReplaceAll(result.Value.Policies);
    }

    private void TryRestoreSelection()
    {
        var currentIdentity = SelectedDevice?.Identity;
        if (string.IsNullOrWhiteSpace(currentIdentity))
        {
            return;
        }

        var match = Devices.FirstOrDefault(device =>
            string.Equals(device.Identity, currentIdentity, StringComparison.OrdinalIgnoreCase));

        SelectedDevice = match;
    }

    private void ApplySelectedDeviceDetails(MobileDeviceListItemDto details)
    {
        var index = Devices
            .Select((device, deviceIndex) => new { device, deviceIndex })
            .FirstOrDefault(entry => string.Equals(entry.device.Identity, details.Identity, StringComparison.OrdinalIgnoreCase))
            ?.deviceIndex;

        if (index.HasValue)
        {
            Devices[index.Value] = details;
        }

        _suppressSelectedDeviceLoad = true;
        try
        {
            SelectedDevice = details;
        }
        finally
        {
            _suppressSelectedDeviceLoad = false;
        }
    }

    private void ResetDisconnectedState()
    {
        CapabilityMessage = null;
        Devices.Clear();
        Policies.Clear();
        ErrorMessage = null;
        TotalCount = 0;
        IsTotalCountExact = true;
        HasMore = false;
        IsLoadingSelection = false;
        ClearLoadingProgress();
        _shellViewModel.ClearPageAlert(AlertPage);
    }

    private void ClearLoadingProgress()
    {
        LoadingProgress = 0;
        LoadingStatus = null;
        LoadingCurrentItem = null;
        LoadingTotalItems = null;
    }
}

