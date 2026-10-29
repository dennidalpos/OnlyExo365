using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Shell.ViewModels;

namespace OnlyExo365.Tests;

public class ResourcesViewModelTests
{
    [Fact]
    public void ResourceMailboxEditorViewModel_BuildPermissionDelta_ComputesCaseInsensitiveAddsAndRemovalsForEveryPermissionType()
    {
        var editor = new ResourceMailboxEditorViewModel();
        editor.ApplyDetails(new ResourceMailboxDetailsDto
        {
            Identity = "room-01",
            DisplayName = "Room 01",
            Alias = "room01",
            PrimarySmtpAddress = "room01@contoso.com",
            Permissions = new MailboxPermissionsDto
            {
                FullAccessPermissions =
                [
                    new MailboxPermissionEntryDto { User = "existing.full@contoso.com" },
                    new MailboxPermissionEntryDto { User = "removed.full@contoso.com" }
                ],
                SendAsPermissions =
                [
                    new RecipientPermissionEntryDto
                    {
                        Trustee = "legacy\\helper",
                        ResolvedTrustee = "helper@contoso.com"
                    },
                    new RecipientPermissionEntryDto
                    {
                        Trustee = "removed.sendas@contoso.com"
                    }
                ],
                SendOnBehalfPermissions =
                [
                    new SendOnBehalfPermissionEntryDto { Identity = "assistant@contoso.com" },
                    new SendOnBehalfPermissionEntryDto { Identity = "removed.sob@contoso.com" }
                ]
            }
        });

        editor.FullAccessUsers = "existing.full@contoso.com; new.full@contoso.com; NEW.FULL@contoso.com";
        editor.SendAsUsers = "helper@contoso.com, new.sendas@contoso.com, helper@contoso.com";
        editor.SendOnBehalfUsers = "assistant@contoso.com; new.sob@contoso.com";

        var actions = editor.BuildPermissionDelta();

        Assert.Equal(6, actions.Count);

        Assert.Contains(actions, action =>
            action.Action == PermissionAction.Add &&
            action.PermissionType == PermissionType.FullAccess &&
            action.User == "new.full@contoso.com" &&
            action.AutoMapping == true);

        Assert.Contains(actions, action =>
            action.Action == PermissionAction.Remove &&
            action.PermissionType == PermissionType.FullAccess &&
            action.User == "removed.full@contoso.com" &&
            action.AutoMapping == null);

        Assert.Contains(actions, action =>
            action.Action == PermissionAction.Add &&
            action.PermissionType == PermissionType.SendAs &&
            action.User == "new.sendas@contoso.com" &&
            action.AutoMapping == null);

        Assert.Contains(actions, action =>
            action.Action == PermissionAction.Remove &&
            action.PermissionType == PermissionType.SendAs &&
            action.User == "removed.sendas@contoso.com");

        Assert.Contains(actions, action =>
            action.Action == PermissionAction.Add &&
            action.PermissionType == PermissionType.SendOnBehalf &&
            action.User == "new.sob@contoso.com");

        Assert.Contains(actions, action =>
            action.Action == PermissionAction.Remove &&
            action.PermissionType == PermissionType.SendOnBehalf &&
            action.User == "removed.sob@contoso.com");
    }

    [Fact]
    public void ResourceMailboxEditorViewModel_BuildPermissionDelta_ReturnsEmptyWhenEditorValuesMatchLoadedPermissions()
    {
        var editor = new ResourceMailboxEditorViewModel();
        editor.ApplyDetails(new ResourceMailboxDetailsDto
        {
            Identity = "room-02",
            DisplayName = "Room 02",
            Alias = "room02",
            PrimarySmtpAddress = "room02@contoso.com",
            Permissions = new MailboxPermissionsDto
            {
                FullAccessPermissions =
                [
                    new MailboxPermissionEntryDto { User = "delegate@contoso.com" }
                ],
                SendAsPermissions =
                [
                    new RecipientPermissionEntryDto
                    {
                        Trustee = "delegate@contoso.com",
                        ResolvedTrustee = "delegate@contoso.com"
                    }
                ],
                SendOnBehalfPermissions =
                [
                    new SendOnBehalfPermissionEntryDto { Identity = "assistant@contoso.com" }
                ]
            }
        });

        var actions = editor.BuildPermissionDelta();

        Assert.Empty(actions);
    }
}

