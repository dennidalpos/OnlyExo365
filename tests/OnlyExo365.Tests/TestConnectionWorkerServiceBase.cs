using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;
using OnlyExo365.Shell.Ipc;

namespace OnlyExo365.Tests;

public abstract class TestConnectionWorkerServiceBase : IConnectionWorkerService
{
    private event EventHandler<WorkerConnectionState>? StateChangedHandlers;
    private event EventHandler<EventEnvelope>? EventReceivedHandlers;
    private event EventHandler<CapabilityMapDto>? CapabilitiesUpdatedHandlers;

    public virtual WorkerConnectionState ConnectionState => WorkerConnectionState.Connected;

    public virtual WorkerStatus Status { get; } = new()
    {
        State = WorkerConnectionState.Connected,
        IsModuleAvailable = true
    };

    public virtual CapabilityMapDto? Capabilities => null;

    public virtual event EventHandler<WorkerConnectionState>? StateChanged
    {
        add => StateChangedHandlers += value;
        remove => StateChangedHandlers -= value;
    }

    public virtual event EventHandler<EventEnvelope>? EventReceived
    {
        add => EventReceivedHandlers += value;
        remove => EventReceivedHandlers -= value;
    }

    public virtual event EventHandler<CapabilityMapDto>? CapabilitiesUpdated
    {
        add => CapabilitiesUpdatedHandlers += value;
        remove => CapabilitiesUpdatedHandlers -= value;
    }

    protected void PublishStateChanged(WorkerConnectionState state)
        => StateChangedHandlers?.Invoke(this, state);

    protected void PublishEventReceived(EventEnvelope envelope)
        => EventReceivedHandlers?.Invoke(this, envelope);

    protected void PublishCapabilitiesUpdated(CapabilityMapDto capabilities)
        => CapabilitiesUpdatedHandlers?.Invoke(this, capabilities);

    public virtual Task<bool> StartWorkerAsync(CancellationToken cancellationToken = default) => Task.FromResult(true);
    public virtual Task StopWorkerAsync() => Task.CompletedTask;
    public virtual Task<bool> RestartWorkerAsync(CancellationToken cancellationToken = default) => Task.FromResult(true);
    public virtual void KillWorker() { }
    public virtual Task<Result<SetWorkerConsoleVisibilityResponse>> SetWorkerConsoleVisibilityAsync(bool isVisible, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<ConnectionStatusDto>> ConnectExchangeAsync(Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> DisconnectExchangeAsync(CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<ConnectionStatusDto>> GetConnectionStatusAsync(CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<CapabilityMapDto>> DetectCapabilitiesAsync(bool forceRefresh = false, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
}

