using ExchangeAdmin.Contracts.Messages;

namespace ExchangeAdmin.Worker.Operations;

public partial class OperationDispatcher
{
    private sealed class SupportOperationsHandler(OperationDispatcher dispatcher) : IOperationAreaHandler
    {
        public IReadOnlyCollection<OperationType> SupportedOperations { get; } =
        [
            OperationType.GetUserLicenses,
            OperationType.SetUserLicense,
            OperationType.GetUsageLocationSuggestion,
            OperationType.SetUserUsageLocation,
            OperationType.GetAvailableLicenses,
            OperationType.CheckPrerequisites,
            OperationType.InstallModule,
            OperationType.SetWorkerConsoleVisibility
        ];

        public Task<ResponseEnvelope> HandleAsync(RequestEnvelope request, CancellationToken cancellationToken)
        {
            var correlationId = request.CorrelationId;

            return request.Operation switch
            {
                OperationType.GetUserLicenses => dispatcher.HandleGetUserLicensesAsync(request, correlationId, cancellationToken),
                OperationType.SetUserLicense => dispatcher.HandleSetUserLicenseAsync(request, correlationId, cancellationToken),
                OperationType.GetUsageLocationSuggestion => dispatcher.HandleGetUsageLocationSuggestionAsync(request, correlationId, cancellationToken),
                OperationType.SetUserUsageLocation => dispatcher.HandleSetUserUsageLocationAsync(request, correlationId, cancellationToken),
                OperationType.GetAvailableLicenses => dispatcher.HandleGetAvailableLicensesAsync(request, correlationId, cancellationToken),
                OperationType.CheckPrerequisites => dispatcher.HandleCheckPrerequisitesAsync(request, correlationId, cancellationToken),
                OperationType.InstallModule => dispatcher.HandleInstallModuleAsync(request, correlationId, cancellationToken),
                OperationType.SetWorkerConsoleVisibility => dispatcher.HandleSetWorkerConsoleVisibilityAsync(request, correlationId, cancellationToken),
                _ => throw new InvalidOperationException($"Unsupported support operation: {request.Operation}")
            };
        }
    }
}
