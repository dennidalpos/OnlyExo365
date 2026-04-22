using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;
using OnlyExo365.Shell.Ipc;
using OnlyExo365.Shell.ViewModels;

namespace OnlyExo365.Tests;

public sealed class MailboxPermissionsEditorViewModelTests
{
    [Fact]
    public async Task LoadAndSaveFolderPermissions_KeepLocalizedCalendarAliasAcrossWorkerCalls()
    {
        var worker = new MailboxPermissionsWorkerService();
        var savingFlag = false;
        string? errorMessage = null;

        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        ErrorDialogService.ConfirmationHandlerOverride = (_, _) => true;
        var viewModel = new MailboxPermissionsEditorViewModel(
            worker,
            shell,
            getMailboxIdentity: () => "mailbox-guid",
            getPrimarySmtpAddress: () => "shared@contoso.com",
            getIsSaving: () => savingFlag,
            setIsSaving: value => savingFlag = value,
            setErrorMessage: value => errorMessage = value,
            refreshAsync: _ => Task.CompletedTask);

        try
        {
            viewModel.FolderPermissionFolderPath = "Calendario";

            await viewModel.LoadFolderPermissionsAsync(CancellationToken.None);

            Assert.Null(errorMessage);
            Assert.Single(worker.GetFolderRequests);
            Assert.Equal("shared@contoso.com", worker.GetFolderRequests[0].MailboxIdentity);
            Assert.Equal("Calendario", worker.GetFolderRequests[0].FolderPath);
            Assert.Equal("shared@contoso.com:\\Calendar", viewModel.FolderPermissionTargetLabel);

            viewModel.NewFolderPermissionUser = "delegate@contoso.com";
            viewModel.NewFolderPermissionRole = "Editor";

            Assert.True(viewModel.AddFolderPermissionCommand.CanExecute(null));
            viewModel.AddFolderPermissionCommand.Execute(null);

            await WaitForConditionAsync(() =>
                worker.SetFolderRequests.Count == 1 &&
                worker.GetFolderRequests.Count == 2);

            Assert.Null(errorMessage);
            Assert.Equal("Calendario", worker.SetFolderRequests[0].FolderPath);
            Assert.Equal(PermissionAction.Add, worker.SetFolderRequests[0].Action);
            Assert.Equal(["Editor"], worker.SetFolderRequests[0].AccessRights);
            Assert.Equal("Calendario", worker.GetFolderRequests[1].FolderPath);
        }
        finally
        {
            ErrorDialogService.ConfirmationHandlerOverride = null;
        }
    }

    private static async Task WaitForConditionAsync(Func<bool> condition)
    {
        for (var attempt = 0; attempt < 20; attempt++)
        {
            if (condition())
            {
                return;
            }

            await Task.Delay(25);
        }

        Assert.True(condition(), "Condition not reached in time.");
    }

    private sealed class MailboxPermissionsWorkerService : TestMailboxesWorkerServiceBase
    {
        public List<GetMailboxFolderPermissionsRequest> GetFolderRequests { get; } = [];
        public List<SetMailboxFolderPermissionRequest> SetFolderRequests { get; } = [];

        public override Task<Result<GetMailboxFolderPermissionsResponse>> GetMailboxFolderPermissionsAsync(
            GetMailboxFolderPermissionsRequest request,
            Action<EventEnvelope>? eventHandler = null,
            CancellationToken cancellationToken = default)
        {
            GetFolderRequests.Add(Clone(request));

            var response = new GetMailboxFolderPermissionsResponse
            {
                MailboxIdentity = request.MailboxIdentity,
                FolderPath = request.FolderPath,
                ResolvedFolderIdentity = $"{request.MailboxIdentity}:\\Calendar"
            };

            foreach (var savedPermission in SetFolderRequests)
            {
                response.Permissions.Add(new MailboxFolderPermissionEntryDto
                {
                    User = savedPermission.User,
                    DisplayName = savedPermission.User,
                    AccessRights = savedPermission.AccessRights.ToList(),
                    IsInherited = false
                });
            }

            return Task.FromResult(Result<GetMailboxFolderPermissionsResponse>.Success(response));
        }

        public override Task<Result> SetMailboxFolderPermissionAsync(
            SetMailboxFolderPermissionRequest request,
            Action<EventEnvelope>? eventHandler = null,
            CancellationToken cancellationToken = default)
        {
            SetFolderRequests.Add(Clone(request));
            return Task.FromResult(Result.Success());
        }

        private static GetMailboxFolderPermissionsRequest Clone(GetMailboxFolderPermissionsRequest request)
            => new()
            {
                MailboxIdentity = request.MailboxIdentity,
                FolderPath = request.FolderPath
            };

        private static SetMailboxFolderPermissionRequest Clone(SetMailboxFolderPermissionRequest request)
            => new()
            {
                MailboxIdentity = request.MailboxIdentity,
                FolderPath = request.FolderPath,
                User = request.User,
                Action = request.Action,
                AccessRights = request.AccessRights.ToList()
            };
    }
}

