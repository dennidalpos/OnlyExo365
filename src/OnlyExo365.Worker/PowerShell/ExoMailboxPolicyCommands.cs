using System.Collections;
using OnlyExo365.Contracts.Dtos;

namespace OnlyExo365.Worker.PowerShell;

internal sealed class ExoMailboxPolicyCommands : ExoCommandModuleBase
{
    private readonly CapabilityDetector _capabilityDetector;

    public ExoMailboxPolicyCommands(PowerShellEngine engine, CapabilityDetector capabilityDetector)
        : base(engine)
    {
        _capabilityDetector = capabilityDetector;
    }

    public async Task<List<RetentionPolicySummaryDto>> GetRetentionPoliciesAsync(
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var script = ExoMailboxScriptFactory.BuildGetRetentionPoliciesScript();

        onLog?.Invoke("Verbose", "Fetching retention policies...");

        var result = await Engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to get retention policies: {result.ErrorMessage}");
        }

        var policies = new List<RetentionPolicySummaryDto>();

        foreach (var output in result.Output)
        {
            if (output.BaseObject is object[] array)
            {
                foreach (var policyHash in array.OfType<Hashtable>())
                {
                    policies.Add(ExoMailboxMapper.ToRetentionPolicySummary(policyHash));
                }

                continue;
            }

            if (output.BaseObject is Hashtable hash)
            {
                policies.Add(ExoMailboxMapper.ToRetentionPolicySummary(hash));
            }
        }

        return policies;
    }

    public async Task SetMailboxSettingsAsync(
        UpdateMailboxSettingsRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var capabilities = await _capabilityDetector.DetectCapabilitiesAsync(cancellationToken: cancellationToken);
        MailboxCasCapabilityGuard.EnsureMailboxSettingsUpdateAvailable(capabilities, request);

        var script = ExoMailboxScriptFactory.BuildSetMailboxSettingsScript(request, capabilities.Features);
        if (string.IsNullOrWhiteSpace(script))
        {
            return;
        }

        onLog?.Invoke("Information", $"Updating mailbox settings for {request.Identity}...");

        var result = await Engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to update mailbox settings: {result.ErrorMessage}");
        }

        onLog?.Invoke("Information", "Mailbox settings updated successfully");
    }

    public async Task SetRetentionPolicyAsync(
        SetRetentionPolicyRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var script = ExoMailboxScriptFactory.BuildSetRetentionPolicyScript(request);

        onLog?.Invoke("Information", $"Updating retention policy for {request.Identity}...");

        var result = await Engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to update retention policy: {result.ErrorMessage}");
        }

        onLog?.Invoke("Information", "Retention policy updated successfully");
    }

    public async Task SetMailboxAutoReplyConfigurationAsync(
        SetMailboxAutoReplyConfigurationRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var script = ExoMailboxScriptFactory.BuildSetMailboxAutoReplyConfigurationScript(request);

        onLog?.Invoke("Information", $"Updating auto-reply configuration for {request.Identity}...");

        var result = await Engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to update auto-reply configuration: {result.ErrorMessage}");
        }

        onLog?.Invoke("Information", "Auto-reply configuration updated successfully");
    }
}

