using OnlyExo365.Contracts.Messages;

namespace OnlyExo365.Worker.Operations;

public partial class OperationDispatcher
{
    private sealed class RecipientOperationsHandler(OperationDispatcher dispatcher) : IOperationAreaHandler
    {
        public IReadOnlyCollection<OperationType> SupportedOperations { get; } =
        [
            OperationType.GetDashboardStats,
            OperationType.GetContacts,
            OperationType.GetContactDetails,
            OperationType.UpsertContact,
            OperationType.RemoveContact,
            OperationType.GetResourceMailboxes,
            OperationType.GetResourceMailboxDetails,
            OperationType.UpsertResourceMailbox,
            OperationType.GetPublicFolders,
            OperationType.GetPublicFolderDetails,
            OperationType.UpsertPublicFolder,
            OperationType.SetPublicFolderClientPermission,
            OperationType.RemovePublicFolder,
            OperationType.GetMobileDevices,
            OperationType.GetMobileDeviceDetails,
            OperationType.GetMobileDeviceMailboxPolicies,
            OperationType.SetMobileDeviceAccessState,
            OperationType.ClearMobileDevice,
            OperationType.SetMobileDeviceMailboxPolicy,
            OperationType.GetMigrationBatches,
            OperationType.GetMigrationEndpoints,
            OperationType.GetMigrationBatchDetails,
            OperationType.UpsertMigrationEndpoint,
            OperationType.TestMigrationEndpoint,
            OperationType.GetMigrationBatchPreflight,
            OperationType.CreateMigrationBatch,
            OperationType.StartMigrationBatch,
            OperationType.CompleteMigrationBatch,
            OperationType.RemoveMigrationBatch,
            OperationType.GetRoleGroups,
            OperationType.GetRoleGroupDetails,
            OperationType.UpsertRoleGroup,
            OperationType.ModifyRoleGroupMember
        ];

        public Task<ResponseEnvelope> HandleAsync(RequestEnvelope request, CancellationToken cancellationToken)
        {
            return request.Operation switch
            {
                OperationType.GetDashboardStats => dispatcher.HandleGetDashboardStatsAsync(request, cancellationToken),
                OperationType.GetContacts => dispatcher.HandleGetContactsAsync(request, cancellationToken),
                OperationType.GetContactDetails => dispatcher.HandleGetContactDetailsAsync(request, cancellationToken),
                OperationType.UpsertContact => dispatcher.HandleUpsertContactAsync(request, cancellationToken),
                OperationType.RemoveContact => dispatcher.HandleRemoveContactAsync(request, cancellationToken),
                OperationType.GetResourceMailboxes => dispatcher.HandleGetResourceMailboxesAsync(request, cancellationToken),
                OperationType.GetResourceMailboxDetails => dispatcher.HandleGetResourceMailboxDetailsAsync(request, cancellationToken),
                OperationType.UpsertResourceMailbox => dispatcher.HandleUpsertResourceMailboxAsync(request, cancellationToken),
                OperationType.GetPublicFolders => dispatcher.HandleGetPublicFoldersAsync(request, cancellationToken),
                OperationType.GetPublicFolderDetails => dispatcher.HandleGetPublicFolderDetailsAsync(request, cancellationToken),
                OperationType.UpsertPublicFolder => dispatcher.HandleUpsertPublicFolderAsync(request, cancellationToken),
                OperationType.SetPublicFolderClientPermission => dispatcher.HandleSetPublicFolderClientPermissionAsync(request, cancellationToken),
                OperationType.RemovePublicFolder => dispatcher.HandleRemovePublicFolderAsync(request, cancellationToken),
                OperationType.GetMobileDevices => dispatcher.HandleGetMobileDevicesAsync(request, cancellationToken),
                OperationType.GetMobileDeviceDetails => dispatcher.HandleGetMobileDeviceDetailsAsync(request, cancellationToken),
                OperationType.GetMobileDeviceMailboxPolicies => dispatcher.HandleGetMobileDeviceMailboxPoliciesAsync(request, cancellationToken),
                OperationType.SetMobileDeviceAccessState => dispatcher.HandleSetMobileDeviceAccessStateAsync(request, cancellationToken),
                OperationType.ClearMobileDevice => dispatcher.HandleClearMobileDeviceAsync(request, cancellationToken),
                OperationType.SetMobileDeviceMailboxPolicy => dispatcher.HandleSetMobileDeviceMailboxPolicyAsync(request, cancellationToken),
                OperationType.GetMigrationBatches => dispatcher.HandleGetMigrationBatchesAsync(request, cancellationToken),
                OperationType.GetMigrationEndpoints => dispatcher.HandleGetMigrationEndpointsAsync(request, cancellationToken),
                OperationType.GetMigrationBatchDetails => dispatcher.HandleGetMigrationBatchDetailsAsync(request, cancellationToken),
                OperationType.UpsertMigrationEndpoint => dispatcher.HandleUpsertMigrationEndpointAsync(request, cancellationToken),
                OperationType.TestMigrationEndpoint => dispatcher.HandleTestMigrationEndpointAsync(request, cancellationToken),
                OperationType.GetMigrationBatchPreflight => dispatcher.HandleGetMigrationBatchPreflightAsync(request, cancellationToken),
                OperationType.CreateMigrationBatch => dispatcher.HandleCreateMigrationBatchAsync(request, cancellationToken),
                OperationType.StartMigrationBatch => dispatcher.HandleStartMigrationBatchAsync(request, cancellationToken),
                OperationType.CompleteMigrationBatch => dispatcher.HandleCompleteMigrationBatchAsync(request, cancellationToken),
                OperationType.RemoveMigrationBatch => dispatcher.HandleRemoveMigrationBatchAsync(request, cancellationToken),
                OperationType.GetRoleGroups => dispatcher.HandleGetRoleGroupsAsync(request, cancellationToken),
                OperationType.GetRoleGroupDetails => dispatcher.HandleGetRoleGroupDetailsAsync(request, cancellationToken),
                OperationType.UpsertRoleGroup => dispatcher.HandleUpsertRoleGroupAsync(request, cancellationToken),
                OperationType.ModifyRoleGroupMember => dispatcher.HandleModifyRoleGroupMemberAsync(request, cancellationToken),
                _ => throw new InvalidOperationException($"Unsupported recipient operation: {request.Operation}")
            };
        }
    }
}

