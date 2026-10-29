using OnlyExo365.Shell.Localization;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;

namespace OnlyExo365.Shell.ViewModels;

public partial class MobileDevicesViewModel
{
    private async Task SetAccessStateAsync(string accessState)
    {
        if (SelectedDevice == null)
        {
            return;
        }

        if (!ConfirmMutation(
                Loc.Get("Presentation.ChangeDeviceAccessState"),
                $"{SelectedDevice.MailboxIdentity} / {SelectedDevice.DeviceId}",
                Loc.GetFormat("MobileDevices.SetStateImpact", accessState),
                Loc.Get("Presentation.ConfirmDeviceStateChange")))
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
                ErrorMessage = result.Error?.Message ?? Loc.GetFormat("MobileDevices.StateFailed", accessState);
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
                Loc.Get("Presentation.RemoteWipeDevice"),
                $"{SelectedDevice.MailboxIdentity} / {SelectedDevice.DeviceId}",
                Loc.Get("Presentation.SendARemoteWipeToTheSelectedDevice"),
                Loc.Get("Presentation.ConfirmRemoteWipe")))
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
                ErrorMessage = result.Error?.Message ?? Loc.Get("Presentation.UnableToSendTheRemoteWipe");
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
                Loc.Get("Presentation.AssignDeviceMailboxPolicy"),
                $"{SelectedDevice.MailboxIdentity} / {SelectedDevice.DeviceId}",
                Loc.GetFormat("MobileDevices.AssignPolicyImpact", label),
                Loc.Get("Presentation.ConfirmPolicyAssignment")))
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
                ErrorMessage = result.Error?.Message ?? Loc.Get("Presentation.UnableToAssignTheMailboxPolicy");
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
            return Loc.Get("Presentation.DefaultTenantBehavior");
        }

        return Policies.FirstOrDefault(policy => string.Equals(policy.Identity, SelectedMailboxPolicyIdentity, StringComparison.OrdinalIgnoreCase))?.Name
            ?? SelectedMailboxPolicyIdentity;
    }
}

