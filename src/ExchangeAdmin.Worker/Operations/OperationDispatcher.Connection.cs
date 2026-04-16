using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Worker.PowerShell;

namespace ExchangeAdmin.Worker.Operations;

public partial class OperationDispatcher
{
    private async Task<ResponseEnvelope> HandleConnectAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Starting Exchange Online connection...");
        ConsoleLogger.Info("Dispatcher", $"Connecting to Exchange Online (correlation: {request.CorrelationId})");

        var result = await _psEngine.ConnectExchangeAsync(
            onVerbose: async (level, msg) => await SendLogAsync(request.CorrelationId, LogLevel.Verbose, msg),
            cancellationToken: cancellationToken);

        ConsoleLogger.Debug("Dispatcher", $"ConnectExchangeAsync completed. Success: {result.Success}, WasCancelled: {result.WasCancelled}");

        if (result.WasCancelled)
        {
            ConsoleLogger.Warning("Dispatcher", "Connection was cancelled");
            return CreateCancelledResponse(request.CorrelationId);
        }

        if (!result.Success)
        {
            ConsoleLogger.Error("Dispatcher", $"Connection failed: {result.ErrorMessage}");
            var (code, isTransient, retryAfter) = result.Errors.Any()
                ? ErrorClassifier.Classify(result.Errors.First())
                : (ErrorCode.AuthenticationFailed, false, (int?)null);

            return CreateErrorResponse(request.CorrelationId, code, result.ErrorMessage ?? "Connection failed", isTransient, retryAfter);
        }

        var (isConnected, upn, org, isGraphConnected, isComplianceConnected) = await _psEngine.GetConnectionStatusAsync(cancellationToken);

        var status = new ConnectionStatusDto
        {
            State = isConnected ? ConnectionState.Connected : ConnectionState.Failed,
            UserPrincipalName = upn,
            Organization = org,
            GraphConnected = isGraphConnected,
            ComplianceConnected = isComplianceConnected,
            ConnectedAt = DateTime.UtcNow
        };

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Connected as {upn} to {org}");

        await SendLogAsync(request.CorrelationId, LogLevel.Verbose, "Detecting capabilities...");
        try
        {
            await _capabilityDetector.DetectCapabilitiesAsync(
                forceRefresh: true,
                onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, LogLevel.Verbose, msg),
                cancellationToken: cancellationToken);
        }
        catch (Exception ex)
        {
            await SendLogAsync(request.CorrelationId, LogLevel.Warning, $"Capability detection failed: {ex.Message}");
        }

        return CreateSuccessResponse(request.CorrelationId, status);
    }

    private async Task<ResponseEnvelope> HandleDisconnectAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Disconnecting from Exchange Online...");

        var result = await _psEngine.DisconnectExchangeAsync(cancellationToken);

        if (result.WasCancelled)
        {
            return CreateCancelledResponse(request.CorrelationId);
        }

        _capabilityDetector.ClearCache();

        var status = new ConnectionStatusDto
        {
            State = ConnectionState.Disconnected,
            GraphConnected = false
        };

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Disconnected from Exchange Online");

        return CreateSuccessResponse(request.CorrelationId, status);
    }

    private async Task<ResponseEnvelope> HandleGetConnectionStatusAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var (isConnected, upn, org, isGraphConnected, isComplianceConnected) = await _psEngine.GetConnectionStatusAsync(cancellationToken);

        var status = new ConnectionStatusDto
        {
            State = isConnected ? ConnectionState.Connected : ConnectionState.Disconnected,
            UserPrincipalName = upn,
            Organization = org,
            GraphConnected = isGraphConnected,
            ComplianceConnected = isComplianceConnected
        };

        return CreateSuccessResponse(request.CorrelationId, status);
    }

    private async Task<ResponseEnvelope> HandleDetectCapabilitiesAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var detectRequest = JsonMessageSerializer.ExtractPayload<DetectCapabilitiesRequest>(request.Payload)
            ?? new DetectCapabilitiesRequest();

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Detecting capabilities...");

        var capabilities = await _capabilityDetector.DetectCapabilitiesAsync(
            forceRefresh: detectRequest.ForceRefresh,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, LogLevel.Verbose, msg),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, capabilities);
    }
}
