using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;

namespace ExchangeAdmin.Presentation.ViewModels;

public partial class MobileDevicesViewModel
{
    private async Task SetAccessStateAsync(string accessState)
    {
        if (SelectedDevice == null)
        {
            return;
        }

        if (!ConfirmMutation(
                "Change device access state",
                $"{SelectedDevice.MailboxIdentity} / {SelectedDevice.DeviceId}",
                $"Set the device status to {accessState}.",
                "Confirm device state change"))
        {
            return;
        }

        IsApplyingAction = true;
        ErrorMessage = null;

        try
        {
            var result = await _workerService.SetMobileDeviceAccessStateAsync(
                new SetMobileDeviceAccessStateRequest
                {
                    MailboxIdentity = SelectedDevice.MailboxIdentity,
                    DeviceId = SelectedDevice.DeviceId,
                    AccessState = accessState
                },
                cancellationToken: CancellationToken.None);

            if (!result.IsSuccess)
            {
                ErrorMessage = result.Error?.Message ?? $"Unable to set status {accessState}.";
                return;
            }

            _shellViewModel.AddLog(LogLevel.Information, $"Device {SelectedDevice.DeviceId} set to {accessState}", "MobileDevices");
            await RefreshAsync(CancellationToken.None);
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
        finally
        {
            IsApplyingAction = false;
        }
    }

    private async Task RemoteWipeAsync(CancellationToken cancellationToken)
    {
        if (SelectedDevice == null)
        {
            return;
        }

        if (!ConfirmMutation(
                "Remote wipe device",
                $"{SelectedDevice.MailboxIdentity} / {SelectedDevice.DeviceId}",
                "Send a remote wipe to the selected device.",
                "Confirm remote wipe"))
        {
            return;
        }

        IsApplyingAction = true;
        ErrorMessage = null;

        try
        {
            var result = await _workerService.ClearMobileDeviceAsync(
                new ClearMobileDeviceRequest
                {
                    Identity = SelectedDevice.Identity
                },
                cancellationToken: cancellationToken);

            if (!result.IsSuccess)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to send the remote wipe.";
                return;
            }

            _shellViewModel.AddLog(LogLevel.Warning, $"Remote wipe sent for {SelectedDevice.DeviceId}", "MobileDevices");
            await RefreshAsync(cancellationToken);
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
        finally
        {
            IsApplyingAction = false;
        }
    }

    private async Task AssignPolicyAsync(CancellationToken cancellationToken)
    {
        if (SelectedDevice == null)
        {
            return;
        }

        var label = GetSelectedPolicyDisplayName();
        if (!ConfirmMutation(
                "Assign device mailbox policy",
                $"{SelectedDevice.MailboxIdentity} / {SelectedDevice.DeviceId}",
                $"Assign the {label} policy to the selected device.",
                "Confirm policy assignment"))
        {
            return;
        }

        IsApplyingAction = true;
        ErrorMessage = null;

        try
        {
            var result = await _workerService.SetMobileDeviceMailboxPolicyAsync(
                new SetMobileDeviceMailboxPolicyRequest
                {
                    MailboxIdentity = SelectedDevice.MailboxIdentity,
                    PolicyIdentity = SelectedMailboxPolicyIdentity
                },
                cancellationToken: cancellationToken);

            if (!result.IsSuccess)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to assign the mailbox policy.";
                return;
            }

            _shellViewModel.AddLog(LogLevel.Information, $"Mailbox policy updated for {SelectedDevice.MailboxIdentity}: {label}", "MobileDevices");
            await RefreshAsync(cancellationToken);
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
        finally
        {
            IsApplyingAction = false;
        }
    }

    private string GetSelectedPolicyDisplayName()
    {
        if (string.IsNullOrWhiteSpace(SelectedMailboxPolicyIdentity))
        {
            return "(default tenant behavior)";
        }

        return Policies.FirstOrDefault(policy => string.Equals(policy.Identity, SelectedMailboxPolicyIdentity, StringComparison.OrdinalIgnoreCase))?.Name
            ?? SelectedMailboxPolicyIdentity;
    }
}
