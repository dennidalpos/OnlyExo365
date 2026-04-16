using ExchangeAdmin.Contracts.Messages;

namespace ExchangeAdmin.Worker.Operations;

public partial class OperationDispatcher
{
    private sealed class MailboxOperationsHandler(OperationDispatcher dispatcher) : IOperationAreaHandler
    {
        public IReadOnlyCollection<OperationType> SupportedOperations { get; } =
        [
            OperationType.GetMailboxProvisioningCandidates,
            OperationType.GetMailboxes,
            OperationType.GetDeletedMailboxes,
            OperationType.GetMailboxDetails,
            OperationType.GetRetentionPolicies,
            OperationType.SetRetentionPolicy,
            OperationType.GetMailboxPermissions,
            OperationType.GetMailboxFolderPermissions,
            OperationType.SetMailboxPermission,
            OperationType.SetMailboxFolderPermission,
            OperationType.ApplyPermissionsDeltaPlan,
            OperationType.UpdateMailboxSettings,
            OperationType.SetMailboxAutoReplyConfiguration,
            OperationType.ConvertMailboxToShared,
            OperationType.ConvertMailboxToRegular,
            OperationType.RestoreMailbox,
            OperationType.GetMailboxSpaceReport,
            OperationType.GetMailboxAccessReport,
            OperationType.CreateMailbox
        ];

        public Task<ResponseEnvelope> HandleAsync(RequestEnvelope request, CancellationToken cancellationToken)
        {
            return request.Operation switch
            {
                OperationType.GetMailboxProvisioningCandidates => dispatcher.HandleGetMailboxProvisioningCandidatesAsync(request, cancellationToken),
                OperationType.GetMailboxes => dispatcher.HandleGetMailboxesAsync(request, cancellationToken),
                OperationType.GetDeletedMailboxes => dispatcher.HandleGetDeletedMailboxesAsync(request, cancellationToken),
                OperationType.GetMailboxDetails => dispatcher.HandleGetMailboxDetailsAsync(request, cancellationToken),
                OperationType.GetRetentionPolicies => dispatcher.HandleGetRetentionPoliciesAsync(request, cancellationToken),
                OperationType.SetRetentionPolicy => dispatcher.HandleSetRetentionPolicyAsync(request, cancellationToken),
                OperationType.GetMailboxPermissions => dispatcher.HandleGetMailboxPermissionsAsync(request, cancellationToken),
                OperationType.GetMailboxFolderPermissions => dispatcher.HandleGetMailboxFolderPermissionsAsync(request, cancellationToken),
                OperationType.SetMailboxPermission => dispatcher.HandleSetMailboxPermissionAsync(request, cancellationToken),
                OperationType.SetMailboxFolderPermission => dispatcher.HandleSetMailboxFolderPermissionAsync(request, cancellationToken),
                OperationType.ApplyPermissionsDeltaPlan => dispatcher.HandleApplyPermissionsDeltaPlanAsync(request, cancellationToken),
                OperationType.UpdateMailboxSettings => dispatcher.HandleUpdateMailboxSettingsAsync(request, cancellationToken),
                OperationType.SetMailboxAutoReplyConfiguration => dispatcher.HandleSetMailboxAutoReplyConfigurationAsync(request, cancellationToken),
                OperationType.ConvertMailboxToShared => dispatcher.HandleConvertMailboxToSharedAsync(request, cancellationToken),
                OperationType.ConvertMailboxToRegular => dispatcher.HandleConvertMailboxToRegularAsync(request, cancellationToken),
                OperationType.RestoreMailbox => dispatcher.HandleRestoreMailboxAsync(request, cancellationToken),
                OperationType.GetMailboxSpaceReport => dispatcher.HandleGetMailboxSpaceReportAsync(request, cancellationToken),
                OperationType.GetMailboxAccessReport => dispatcher.HandleGetMailboxAccessReportAsync(request, cancellationToken),
                OperationType.CreateMailbox => dispatcher.HandleCreateMailboxAsync(request, cancellationToken),
                _ => throw new InvalidOperationException($"Unsupported mailbox operation: {request.Operation}")
            };
        }
    }
}
