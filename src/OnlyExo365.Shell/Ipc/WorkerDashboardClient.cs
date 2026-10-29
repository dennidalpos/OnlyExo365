using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;

namespace OnlyExo365.Shell.Ipc;

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

