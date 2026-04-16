using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;

namespace ExchangeAdmin.Infrastructure.Ipc;

internal sealed class WorkerDashboardClient
{
    private readonly WorkerClientRuntime _runtime;

    public WorkerDashboardClient(WorkerClientRuntime runtime)
    {
        _runtime = runtime;
    }

    public Task<Result<DashboardStatsDto>> GetDashboardStatsAsync(
        GetDashboardStatsRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<DashboardStatsDto>(
            OperationType.GetDashboardStats,
            request ?? new GetDashboardStatsRequest(),
            eventHandler,
            cancellationToken);
}
