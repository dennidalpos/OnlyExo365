using OnlyExo365.Contracts.Messages;

namespace OnlyExo365.Worker.Operations;

public partial class OperationDispatcher
{
    private sealed class MailSecurityOperationsHandler(OperationDispatcher dispatcher) : IOperationAreaHandler
    {
        public IReadOnlyCollection<OperationType> SupportedOperations { get; } =
        [
            OperationType.GetMailSecurityBaseline,
            OperationType.UpdateDkimSigningConfig,
            OperationType.UpdateHostedContentFilterPolicy,
            OperationType.UpdateAntiPhishPolicy,
            OperationType.UpdateMalwareFilterPolicy,
            OperationType.UpdateQuarantinePolicy,
            OperationType.UpdateHostedOutboundSpamFilterPolicy
        ];

        public Task<ResponseEnvelope> HandleAsync(RequestEnvelope request, CancellationToken cancellationToken)
        {
            var correlationId = request.CorrelationId;

            return request.Operation switch
            {
                OperationType.GetMailSecurityBaseline => dispatcher.HandleGetMailSecurityBaselineAsync(request, correlationId, cancellationToken),
                OperationType.UpdateDkimSigningConfig => dispatcher.HandleUpdateDkimSigningConfigAsync(request, correlationId, cancellationToken),
                OperationType.UpdateHostedContentFilterPolicy => dispatcher.HandleUpdateHostedContentFilterPolicyAsync(request, correlationId, cancellationToken),
                OperationType.UpdateAntiPhishPolicy => dispatcher.HandleUpdateAntiPhishPolicyAsync(request, correlationId, cancellationToken),
                OperationType.UpdateMalwareFilterPolicy => dispatcher.HandleUpdateMalwareFilterPolicyAsync(request, correlationId, cancellationToken),
                OperationType.UpdateQuarantinePolicy => dispatcher.HandleUpdateQuarantinePolicyAsync(request, correlationId, cancellationToken),
                OperationType.UpdateHostedOutboundSpamFilterPolicy => dispatcher.HandleUpdateHostedOutboundSpamFilterPolicyAsync(request, correlationId, cancellationToken),
                _ => throw new InvalidOperationException($"Unsupported mail security operation: {request.Operation}")
            };
        }
    }
}

