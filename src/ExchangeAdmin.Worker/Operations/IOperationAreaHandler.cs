using ExchangeAdmin.Contracts.Messages;

namespace ExchangeAdmin.Worker.Operations;

internal interface IOperationAreaHandler
{
    IReadOnlyCollection<OperationType> SupportedOperations { get; }

    Task<ResponseEnvelope> HandleAsync(RequestEnvelope request, CancellationToken cancellationToken);
}
