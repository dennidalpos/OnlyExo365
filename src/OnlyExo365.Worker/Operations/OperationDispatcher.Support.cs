using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Worker.PowerShell;

namespace OnlyExo365.Worker.Operations;

public partial class OperationDispatcher
{
    private async Task<ResponseEnvelope> HandleGetUserLicensesAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var licenseRequest = JsonMessageSerializer.ExtractPayload<GetUserLicensesRequest>(request.Payload);

        if (licenseRequest == null || string.IsNullOrWhiteSpace(licenseRequest.UserPrincipalName))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "UserPrincipalName is required");
        }

        await SendLogAsync(correlationId, LogLevel.Information, $"Fetching licenses for {licenseRequest.UserPrincipalName}...");
        await SendProgressAsync(correlationId, 0, "Fetching user licenses...");

        var response = await _exoCommands.GetUserLicensesAsync(
            licenseRequest.UserPrincipalName,
            cancellationToken);

        await SendProgressAsync(correlationId, 100, "User licenses retrieved");

        return CreateSuccessResponse(correlationId, response);
    }

    private async Task<ResponseEnvelope> HandleSetUserLicenseAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var setRequest = JsonMessageSerializer.ExtractPayload<SetUserLicenseRequest>(request.Payload);

        if (setRequest == null || string.IsNullOrWhiteSpace(setRequest.UserPrincipalName))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "UserPrincipalName is required");
        }

        await SendLogAsync(correlationId, LogLevel.Information, $"Updating licenses for {setRequest.UserPrincipalName}...");
        await SendProgressAsync(correlationId, 0, "Setting user license...");

        await _exoCommands.SetUserLicenseAsync(
            setRequest,
            cancellationToken);

        await SendProgressAsync(correlationId, 100, "User license updated");

        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleGetUsageLocationSuggestionAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var suggestionRequest = JsonMessageSerializer.ExtractPayload<GetUsageLocationSuggestionRequest>(request.Payload);

        if (suggestionRequest == null || string.IsNullOrWhiteSpace(suggestionRequest.UserPrincipalName))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "UserPrincipalName is required");
        }

        await SendLogAsync(correlationId, LogLevel.Information, $"Resolving usage location suggestion for {suggestionRequest.UserPrincipalName}...");
        await SendProgressAsync(correlationId, 0, "Resolving usage location suggestion...");

        var response = await _exoCommands.GetUsageLocationSuggestionAsync(suggestionRequest, cancellationToken);

        await SendProgressAsync(correlationId, 100, "Usage location suggestion resolved");

        return CreateSuccessResponse(correlationId, response);
    }

    private async Task<ResponseEnvelope> HandleSetUserUsageLocationAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var setRequest = JsonMessageSerializer.ExtractPayload<SetUserUsageLocationRequest>(request.Payload);

        if (setRequest == null || string.IsNullOrWhiteSpace(setRequest.UserPrincipalName) || string.IsNullOrWhiteSpace(setRequest.UsageLocation))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "UserPrincipalName and UsageLocation are required");
        }

        await SendLogAsync(correlationId, LogLevel.Information, $"Updating usage location for {setRequest.UserPrincipalName}...");
        await SendProgressAsync(correlationId, 0, "Setting user usage location...");

        await _exoCommands.SetUserUsageLocationAsync(setRequest, cancellationToken);

        await SendProgressAsync(correlationId, 100, "User usage location updated");

        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleGetAvailableLicensesAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        await SendLogAsync(correlationId, LogLevel.Information, "Fetching available licenses...");
        await SendProgressAsync(correlationId, 0, "Fetching available licenses...");

        var licenses = await _exoCommands.GetTenantLicensesAsync(cancellationToken);

        var response = new GetAvailableLicensesResponse
        {
            Licenses = licenses
        };

        await SendProgressAsync(correlationId, 100, "Available licenses retrieved");

        return CreateSuccessResponse(correlationId, response);
    }

    private async Task<ResponseEnvelope> HandleCheckPrerequisitesAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        await SendLogAsync(correlationId, LogLevel.Information, "Checking prerequisites...");
        await SendProgressAsync(correlationId, 0, "Checking system prerequisites...");

        await SendLogAsync(correlationId, LogLevel.Information, "[Prerequisites] Running PowerShell/module checks");
        var status = await _exoCommands.CheckPrerequisitesAsync(cancellationToken);

        await SendProgressAsync(correlationId, 100, "Prerequisite check complete");

        return CreateSuccessResponse(correlationId, status);
    }

    private async Task<ResponseEnvelope> HandleInstallModuleAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var installRequest = JsonMessageSerializer.ExtractPayload<InstallModuleRequest>(request.Payload);

        if (installRequest == null || string.IsNullOrWhiteSpace(installRequest.ModuleName))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "ModuleName is required");
        }

        var installLabel = string.Equals(installRequest.InstallTarget, "PowerShell7", StringComparison.OrdinalIgnoreCase)
            ? "PowerShell 7"
            : installRequest.ModuleName;

        await SendLogAsync(correlationId, LogLevel.Information, $"Installing {installLabel}...");
        await SendProgressAsync(correlationId, 0, $"Installing {installLabel}...");

        await SendLogAsync(correlationId, LogLevel.Information, $"[ModuleInstall] Starting install: {installLabel}");
        var response = await _exoCommands.InstallModuleAsync(
            installRequest,
            cancellationToken);

        await SendProgressAsync(correlationId, 100, $"{installLabel} installation complete");

        return CreateSuccessResponse(correlationId, response);
    }

    private async Task<ResponseEnvelope> HandleSetWorkerConsoleVisibilityAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var visibilityRequest = JsonMessageSerializer.ExtractPayload<SetWorkerConsoleVisibilityRequest>(request.Payload);
        if (visibilityRequest == null)
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "IsVisible is required");
        }

        await SendLogAsync(
            correlationId,
            LogLevel.Information,
            visibilityRequest.IsVisible ? "Showing worker console..." : "Hiding worker console...");

        var response = _consoleController.SetVisibility(visibilityRequest.IsVisible);

        await SendLogAsync(correlationId, LogLevel.Information, response.Message ?? "Worker console visibility updated.");

        return CreateSuccessResponse(correlationId, response);
    }

}

