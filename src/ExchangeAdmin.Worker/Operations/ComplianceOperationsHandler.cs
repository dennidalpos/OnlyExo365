using ExchangeAdmin.Contracts.Messages;

namespace ExchangeAdmin.Worker.Operations;

public partial class OperationDispatcher
{
    private sealed class ComplianceOperationsHandler(OperationDispatcher dispatcher) : IOperationAreaHandler
    {
        public IReadOnlyCollection<OperationType> SupportedOperations { get; } =
        [
            OperationType.GetComplianceWorkspace,
            OperationType.SearchUnifiedAuditLog,
            OperationType.CreateComplianceSearch,
            OperationType.StartComplianceSearch,
            OperationType.RemoveComplianceSearch,
            OperationType.InvokeComplianceAction
        ];

        public Task<ResponseEnvelope> HandleAsync(RequestEnvelope request, CancellationToken cancellationToken)
        {
            return request.Operation switch
            {
                OperationType.GetComplianceWorkspace => dispatcher.HandleGetComplianceWorkspaceAsync(request, cancellationToken),
                OperationType.SearchUnifiedAuditLog => dispatcher.HandleSearchUnifiedAuditLogAsync(request, cancellationToken),
                OperationType.CreateComplianceSearch => dispatcher.HandleCreateComplianceSearchAsync(request, cancellationToken),
                OperationType.StartComplianceSearch => dispatcher.HandleStartComplianceSearchAsync(request, cancellationToken),
                OperationType.RemoveComplianceSearch => dispatcher.HandleRemoveComplianceSearchAsync(request, cancellationToken),
                OperationType.InvokeComplianceAction => dispatcher.HandleInvokeComplianceActionAsync(request, cancellationToken),
                _ => throw new InvalidOperationException($"Unsupported compliance operation: {request.Operation}")
            };
        }
    }
}
