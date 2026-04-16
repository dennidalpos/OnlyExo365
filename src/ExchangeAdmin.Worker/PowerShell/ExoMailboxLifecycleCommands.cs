using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

internal sealed class ExoMailboxLifecycleCommands : ExoCommandModuleBase
{
    public ExoMailboxLifecycleCommands(PowerShellEngine engine)
        : base(engine)
    {
    }

    public async Task CreateMailboxAsync(
        CreateMailboxRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var command = ExoMailboxScriptFactory.BuildCreateMailboxCommand(request);

        onLog?.Invoke("Information", $"Creating mailbox {request.PrimarySmtpAddress}...");

        var result = await Engine.ExecuteAsync(command.Script, command.Parameters, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create mailbox: {result.ErrorMessage}");
        }

        onLog?.Invoke("Information", "Mailbox created successfully");
    }

    public Task ConvertMailboxToSharedAsync(
        ConvertMailboxToSharedRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
        => ConvertMailboxTypeAsync(request.Identity, "Shared", "shared", onLog, cancellationToken);

    public Task ConvertMailboxToRegularAsync(
        ConvertMailboxToRegularRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
        => ConvertMailboxTypeAsync(request.Identity, "Regular", "regular", onLog, cancellationToken);

    public async Task<RestoreMailboxResponse> RestoreMailboxAsync(
        RestoreMailboxRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var command = ExoMailboxScriptFactory.BuildRestoreMailboxCommand(request);

        onLog?.Invoke("Information", $"Starting mailbox restore for {request.SourceIdentity}...");

        var result = await Engine.ExecuteAsync(
            command.Script,
            command.Parameters,
            onVerbose: onLog,
            cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success && result.Output.Count == 0)
        {
            throw new InvalidOperationException($"Failed to restore mailbox: {result.ErrorMessage}");
        }

        return ExoMailboxMapper.ToRestoreMailboxResponse(result, request);
    }

    private async Task ConvertMailboxTypeAsync(
        string identity,
        string mailboxType,
        string friendlyMailboxType,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var script = ExoMailboxScriptFactory.BuildConvertMailboxTypeScript(identity, mailboxType);

        onLog?.Invoke("Information", $"Converting mailbox {identity} to {friendlyMailboxType}...");

        var result = await Engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to convert mailbox: {result.ErrorMessage}");
        }

        onLog?.Invoke("Information", $"Mailbox converted to {friendlyMailboxType} successfully");
    }
}
