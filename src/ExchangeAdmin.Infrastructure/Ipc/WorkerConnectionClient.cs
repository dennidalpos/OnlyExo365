using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;

namespace ExchangeAdmin.Infrastructure.Ipc;

internal sealed class WorkerConnectionClient
{
    private readonly WorkerClientRuntime _runtime;
    private readonly Func<CapabilityMapDto?> _getCapabilities;
    private readonly Action<CapabilityMapDto?> _setCapabilities;
    private readonly Action<CapabilityMapDto> _notifyCapabilitiesUpdated;

    public WorkerConnectionClient(
        WorkerClientRuntime runtime,
        Func<CapabilityMapDto?> getCapabilities,
        Action<CapabilityMapDto?> setCapabilities,
        Action<CapabilityMapDto> notifyCapabilitiesUpdated)
    {
        _runtime = runtime;
        _getCapabilities = getCapabilities;
        _setCapabilities = setCapabilities;
        _notifyCapabilitiesUpdated = notifyCapabilitiesUpdated;
    }

    public async Task<Result<ConnectionStatusDto>> ConnectExchangeAsync(
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        var result = await _runtime.ExecuteOperationAsync<ConnectionStatusDto>(
            OperationType.ConnectExchangeInteractive,
            null,
            eventHandler,
            cancellationToken);

        if (result.IsSuccess && result.Value?.State == ConnectionState.Connected)
        {
            _ = DetectCapabilitiesAsync(forceRefresh: true, cancellationToken: CancellationToken.None);
        }

        return result;
    }

    public async Task<Result> DisconnectExchangeAsync(CancellationToken cancellationToken = default)
    {
        var result = await _runtime.ExecuteCommandAsync(
            OperationType.DisconnectExchange,
            null,
            null,
            cancellationToken);

        _setCapabilities(null);
        return result;
    }

    public Task<Result<ConnectionStatusDto>> GetConnectionStatusAsync(CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<ConnectionStatusDto>(
            OperationType.GetConnectionStatus,
            null,
            null,
            cancellationToken);

    public async Task<Result<CapabilityMapDto>> DetectCapabilitiesAsync(
        bool forceRefresh = false,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        if (!forceRefresh && _getCapabilities() is { } cachedCapabilities)
            return Result<CapabilityMapDto>.Success(cachedCapabilities);

        var result = await _runtime.ExecuteOperationAsync<CapabilityMapDto>(
            OperationType.DetectCapabilities,
            new DetectCapabilitiesRequest { ForceRefresh = forceRefresh },
            eventHandler,
            cancellationToken);

        if (result.IsSuccess && result.Value != null)
        {
            _setCapabilities(result.Value);
            _notifyCapabilitiesUpdated(result.Value);
        }

        return result;
    }
}
