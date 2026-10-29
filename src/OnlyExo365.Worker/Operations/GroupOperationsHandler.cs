using OnlyExo365.Contracts.Messages;

namespace OnlyExo365.Worker.Operations;

public partial class OperationDispatcher
{
    private sealed class GroupOperationsHandler(OperationDispatcher dispatcher) : IOperationAreaHandler
    {
        public IReadOnlyCollection<OperationType> SupportedOperations { get; } =
        [
            OperationType.GetDistributionLists,
            OperationType.GetDistributionListDetails,
            OperationType.GetGroupMembers,
            OperationType.ModifyGroupMember,
            OperationType.PreviewDynamicGroupMembers,
            OperationType.SetDistributionListSettings,
            OperationType.CreateDistributionList
        ];

        public Task<ResponseEnvelope> HandleAsync(RequestEnvelope request, CancellationToken cancellationToken)
        {
            return request.Operation switch
            {
                OperationType.GetDistributionLists => dispatcher.HandleGetDistributionListsAsync(request, cancellationToken),
                OperationType.GetDistributionListDetails => dispatcher.HandleGetDistributionListDetailsAsync(request, cancellationToken),
                OperationType.GetGroupMembers => dispatcher.HandleGetGroupMembersAsync(request, cancellationToken),
                OperationType.ModifyGroupMember => dispatcher.HandleModifyGroupMemberAsync(request, cancellationToken),
                OperationType.PreviewDynamicGroupMembers => dispatcher.HandlePreviewDynamicGroupMembersAsync(request, cancellationToken),
                OperationType.SetDistributionListSettings => dispatcher.HandleSetDistributionListSettingsAsync(request, cancellationToken),
                OperationType.CreateDistributionList => dispatcher.HandleCreateDistributionListAsync(request, cancellationToken),
                _ => throw new InvalidOperationException($"Unsupported group operation: {request.Operation}")
            };
        }
    }
}

