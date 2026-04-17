namespace ExchangeAdmin.Tests;

public sealed class PresentationViewModelCharacterizationTests
{
    [Fact]
    public void PermissionsViewModel_UsesSharedPagingDebounceAndMutationConfirmation()
    {
        var content = ReadViewModel("PermissionsViewModel.cs");

        Assert.Contains("PagingDefaults.DefaultPageSize", content, StringComparison.Ordinal);
        Assert.Contains("_searchDebounce.Debounce(TriggerRefreshFromUi, 300);", content, StringComparison.Ordinal);
        Assert.Contains("ConfirmMutation(", content, StringComparison.Ordinal);
        Assert.Contains("GetRoleGroupsAsync", content, StringComparison.Ordinal);
        Assert.Contains("GetRoleGroupDetailsAsync", content, StringComparison.Ordinal);
        Assert.Contains("ModifyRoleGroupMemberAsync", content, StringComparison.Ordinal);
        Assert.Contains("SelectedRoleGroupSummary", content, StringComparison.Ordinal);
        Assert.Contains("ScopesText", content, StringComparison.Ordinal);
    }

    [Fact]
    public void PublicFoldersViewModel_TracksDirtyStatePermissionRolesAndRecursiveDelete()
    {
        var content = ReadViewModel("PublicFoldersViewModel.cs");

        Assert.Contains("PublicFolderPermissionRoles", content, StringComparison.Ordinal);
        Assert.Contains("NormalizeParentPath", content, StringComparison.Ordinal);
        Assert.Contains("SetTrackedProperty", content, StringComparison.Ordinal);
        Assert.Contains("HasPendingChanges = true;", content, StringComparison.Ordinal);
        Assert.Contains("SetPublicFolderClientPermissionAsync", content, StringComparison.Ordinal);
        Assert.Contains("Recursive = HasSubFolders", content, StringComparison.Ordinal);
        Assert.Contains("Confirm public folder permission update", content, StringComparison.Ordinal);
    }

    [Fact]
    public void LogsViewModel_UsesPersistentLogExportAndDebouncedRefresh()
    {
        var content = ReadViewModel("LogsViewModel.cs");

        Assert.Contains("PersistentLogWriter.GetDefaultLogDirectoryPath()", content, StringComparison.Ordinal);
        Assert.Contains("PersistentLogStore.ApplyRetention(", content, StringComparison.Ordinal);
        Assert.Contains("PersistentLogStore.ExportLogArchive", content, StringComparison.Ordinal);
        Assert.Contains("ExcelExportService.ResolveExportDirectory()", content, StringComparison.Ordinal);
        Assert.Contains("SaveFileDialog", content, StringComparison.Ordinal);
        Assert.Contains("MinRefreshInterval = TimeSpan.FromMilliseconds(100)", content, StringComparison.Ordinal);
        Assert.Contains("_refreshDebounce.Debounce", content, StringComparison.Ordinal);
        Assert.Contains("PersistentObservabilitySummaryPath", content, StringComparison.Ordinal);
        Assert.DoesNotContain("private static string ResolveExportDirectory()", content, StringComparison.Ordinal);
    }

    [Fact]
    public void ShellPromptViewModel_UsesDispatcherFrameForBlockingConfirmation()
    {
        var content = ReadViewModel("ShellPromptViewModel.cs");

        Assert.Contains("TaskCompletionSource<bool>", content, StringComparison.Ordinal);
        Assert.Contains("public bool ShowConfirmationBlocking", content, StringComparison.Ordinal);
        Assert.Contains("DispatcherFrame", content, StringComparison.Ordinal);
        Assert.Contains("Dispatcher.PushFrame(frame);", content, StringComparison.Ordinal);
        Assert.Contains("ResolvePendingConfirmation(true);", content, StringComparison.Ordinal);
        Assert.Contains("ResolvePendingConfirmation(false);", content, StringComparison.Ordinal);
        Assert.Contains("ConfirmCommand = new RelayCommand(Confirm", content, StringComparison.Ordinal);
    }

    [Fact]
    public void WorkspaceViewModels_DoNotUseLocalDisconnectedExchangeBannerCopy()
    {
        string[] files =
        [
            "ComplianceViewModel.cs",
            "DeletedMailboxesViewModel.cs",
            "MailboxAccessReportViewModel.cs",
            "MailboxListViewModel.cs",
            "MailboxProvisioningCandidatesViewModel.cs",
            "MailboxSpaceViewModel.cs",
            "MailFlowViewModel.cs",
            "MessageTraceViewModel.cs",
            "MigrationViewModel.cs",
            "MigrationViewModel.EndpointEditor.cs",
            "MobileDevicesViewModel.cs",
            "MobileDevicesViewModel.Loading.cs",
            "PermissionsViewModel.cs",
            "PublicFoldersViewModel.cs"
        ];

        foreach (var file in files)
        {
            var content = ReadViewModel(file);
            Assert.DoesNotContain("Not connected to Exchange Online", content, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void RemainingComposedWorkspaceViewModels_PublishBlockingFirstLoadFailuresThroughShellAlerts()
    {
        string[] files =
        [
            "ComplianceViewModel.cs",
            "MailFlowViewModel.cs",
            "MigrationViewModel.cs",
            "MigrationViewModel.EndpointEditor.cs",
            "MobileDevicesViewModel.Loading.cs"
        ];

        foreach (var file in files)
        {
            var content = ReadViewModel(file);
            Assert.Contains("ShowPageLoadFailedAlert", content, StringComparison.Ordinal);
            Assert.Contains("ClearPageAlert", content, StringComparison.Ordinal);
        }
    }

    private static string ReadViewModel(string fileName)
        => File.ReadAllText(TestPathHelper.GetRepositoryPath("src", "ExchangeAdmin.Presentation", "ViewModels", fileName));
}
