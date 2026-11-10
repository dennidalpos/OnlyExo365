using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;

namespace OnlyExo365.Shell.Ipc;

internal sealed class WorkerSupportClient
{
    private readonly WorkerClientRuntime _runtime;

    public WorkerSupportClient(WorkerClientRuntime runtime)
    {
        _runtime = runtime;
    }

    public Task<Result<GetUserLicensesResponse>> GetUserLicensesAsync(
        GetUserLicensesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetUserLicensesResponse>(OperationType.GetUserLicenses, request, eventHandler, cancellationToken);

    public Task<Result> SetUserLicenseAsync(
        SetUserLicenseRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.SetUserLicense, request, eventHandler, cancellationToken);

    public Task<Result<GetUsageLocationSuggestionResponse>> GetUsageLocationSuggestionAsync(
        GetUsageLocationSuggestionRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetUsageLocationSuggestionResponse>(OperationType.GetUsageLocationSuggestion, request, eventHandler, cancellationToken);

    public Task<Result> SetUserUsageLocationAsync(
        SetUserUsageLocationRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.SetUserUsageLocation, request, eventHandler, cancellationToken);

    public Task<Result<GetAvailableLicensesResponse>> GetAvailableLicensesAsync(
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetAvailableLicensesResponse>(OperationType.GetAvailableLicenses, null, eventHandler, cancellationToken);

    public Task<Result<PrerequisiteStatusDto>> CheckPrerequisitesAsync(
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<PrerequisiteStatusDto>(OperationType.CheckPrerequisites, null, eventHandler, cancellationToken);

    public Task<Result<InstallModuleResponse>> InstallModuleAsync(
        InstallModuleRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<InstallModuleResponse>(OperationType.InstallModule, request, eventHandler, cancellationToken);

    public async Task<Result<SetWorkerConsoleVisibilityResponse>> SetWorkerConsoleVisibilityAsync(
        bool isVisible,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        var result = await _runtime.ExecuteOperationAsync<SetWorkerConsoleVisibilityResponse>(
            OperationType.SetWorkerConsoleVisibility,
            new SetWorkerConsoleVisibilityRequest { IsVisible = isVisible },
            eventHandler,
            cancellationToken).ConfigureAwait(false);

        if (result.IsSuccess && result.Value != null)
        {
            _runtime.UpdateWorkerConsoleVisibility(result.Value.IsVisible);
        }

        return result;
    }

}

