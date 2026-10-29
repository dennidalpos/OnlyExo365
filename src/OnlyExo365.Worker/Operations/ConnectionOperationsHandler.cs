using OnlyExo365.Contracts.Messages;

namespace OnlyExo365.Worker.Operations;

public partial class OperationDispatcher
{
    private sealed class ConnectionOperationsHandler(OperationDispatcher dispatcher) : IOperationAreaHandler
    {
        public IReadOnlyCollection<OperationType> SupportedOperations { get; } =
        [
            OperationType.ConnectExchangeInteractive,
            OperationType.DisconnectExchange,
            OperationType.GetConnectionStatus,
            OperationType.DetectCapabilities
        ];

        public Task<ResponseEnvelope> HandleAsync(RequestEnvelope request, CancellationToken cancellationToken)
        {
            return request.Operation switch
            {
                OperationType.ConnectExchangeInteractive => dispatcher.HandleConnectAsync(request, cancellationToken),
                OperationType.DisconnectExchange => dispatcher.HandleDisconnectAsync(request, cancellationToken),
                OperationType.GetConnectionStatus => dispatcher.HandleGetConnectionStatusAsync(request, cancellationToken),
                OperationType.DetectCapabilities => dispatcher.HandleDetectCapabilitiesAsync(request, cancellationToken),
                _ => throw new InvalidOperationException($"Unsupported connection operation: {request.Operation}")
            };
        }
    }
}

