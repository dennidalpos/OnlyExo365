using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;

namespace ExchangeAdmin.Tests;

public abstract class TestMigrationWorkerServiceBase : IMigrationWorkerService
{
    public virtual Task<Result<GetMigrationBatchesResponse>> GetMigrationBatchesAsync(GetMigrationBatchesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetMigrationEndpointsResponse>> GetMigrationEndpointsAsync(GetMigrationEndpointsRequest? request = null, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<MigrationBatchDetailsDto>> GetMigrationBatchDetailsAsync(GetMigrationBatchDetailsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpsertMigrationEndpointAsync(UpsertMigrationEndpointRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<TestMigrationEndpointResponse>> TestMigrationEndpointAsync(TestMigrationEndpointRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetMigrationBatchPreflightResponse>> GetMigrationBatchPreflightAsync(GetMigrationBatchPreflightRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> CreateMigrationBatchAsync(CreateMigrationBatchRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> StartMigrationBatchAsync(StartMigrationBatchRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> CompleteMigrationBatchAsync(CompleteMigrationBatchRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> RemoveMigrationBatchAsync(RemoveMigrationBatchRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
}
