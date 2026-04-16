using System.Collections;
using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

internal sealed class ExoMailboxDetailCommands : ExoCommandModuleBase
{
    private readonly CapabilityDetector _capabilityDetector;
    private readonly ExoMailboxReportingCommands _mailboxReportingCommands;
    private readonly ExoMailboxLicenseCommands _mailboxLicenseCommands;

    public ExoMailboxDetailCommands(
        PowerShellEngine engine,
        CapabilityDetector capabilityDetector,
        ExoMailboxReportingCommands mailboxReportingCommands,
        ExoMailboxLicenseCommands mailboxLicenseCommands)
        : base(engine)
    {
        _capabilityDetector = capabilityDetector;
        _mailboxReportingCommands = mailboxReportingCommands;
        _mailboxLicenseCommands = mailboxLicenseCommands;
    }

    public async Task<MailboxDetailsDto> GetMailboxDetailsAsync(
        GetMailboxDetailsRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var script = ExoMailboxScriptFactory.BuildGetMailboxDetailsScript(request.Identity);

        onLog?.Invoke("Verbose", $"Fetching mailbox details for {request.Identity}...");

        var result = await Engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success || !result.Output.Any())
        {
            var reason = string.IsNullOrWhiteSpace(result.ErrorMessage)
                ? $"No output returned for identity '{request.Identity}'"
                : result.ErrorMessage;
            throw new InvalidOperationException($"Failed to get mailbox: {reason}");
        }

        if (result.Output.First().BaseObject is not Hashtable hash)
        {
            throw new InvalidOperationException("Failed to parse mailbox data");
        }

        var details = ExoMailboxMapper.ToMailboxDetails(hash);

        cancellationToken.ThrowIfCancellationRequested();

        if (request.IncludeStatistics)
        {
            details.Statistics = await GetMailboxStatisticsAsync(request.Identity, onLog, cancellationToken);
        }

        cancellationToken.ThrowIfCancellationRequested();

        if (request.IncludeRules)
        {
            details.InboxRules = await GetInboxRulesAsync(request.Identity, onLog, cancellationToken);
        }

        cancellationToken.ThrowIfCancellationRequested();

        if (request.IncludeAutoReply)
        {
            details.AutoReplyConfiguration = await GetAutoReplyConfigurationAsync(request.Identity, onLog, cancellationToken);
        }

        cancellationToken.ThrowIfCancellationRequested();

        if (request.IncludePermissions)
        {
            details.Permissions = await _mailboxReportingCommands.GetMailboxPermissionsAsync(request.Identity, onLog, cancellationToken);
        }

        cancellationToken.ThrowIfCancellationRequested();

        await TryEnrichCasMailboxSettingsAsync(details, onLog, cancellationToken);

        cancellationToken.ThrowIfCancellationRequested();

        if (!string.IsNullOrEmpty(details.UserPrincipalName))
        {
            try
            {
                onLog?.Invoke("Verbose", $"Fetching user licenses for {details.UserPrincipalName}...");
                var licenseResponse = await _mailboxLicenseCommands.GetUserLicensesAsync(details.UserPrincipalName, cancellationToken);
                details.AssignedLicenses = licenseResponse.Licenses;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                onLog?.Invoke("Warning", $"Could not fetch user licenses: {ex.Message}");
            }
        }

        onLog?.Invoke("Information", $"Retrieved details for mailbox {details.DisplayName}");
        return details;
    }

    private async Task TryEnrichCasMailboxSettingsAsync(
        MailboxDetailsDto details,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var capabilities = await _capabilityDetector.DetectCapabilitiesAsync(cancellationToken: cancellationToken);
        if (!capabilities.Features.CanGetCasMailbox)
        {
            onLog?.Invoke("Verbose", "Skipping CAS mailbox settings: Get-CASMailbox not available.");
            return;
        }

        var script = ExoMailboxScriptFactory.BuildGetMailboxCasSettingsScript(details.Identity);
        onLog?.Invoke("Verbose", "Fetching CAS mailbox settings...");

        try
        {
            var result = await Engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
            if (result.WasCancelled)
            {
                throw new OperationCanceledException();
            }

            if (result.Success && result.Output.FirstOrDefault()?.BaseObject is Hashtable hash)
            {
                ExoMailboxMapper.ApplyCasMailboxSettings(details.Features, hash);
                return;
            }

            if (!result.Success)
            {
                onLog?.Invoke("Warning", $"Could not fetch CAS mailbox settings: {result.ErrorMessage}");
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            onLog?.Invoke("Warning", $"Could not fetch CAS mailbox settings: {ex.Message}");
        }
    }

    private async Task<MailboxStatisticsDto?> GetMailboxStatisticsAsync(
        string identity,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var script = ExoMailboxScriptFactory.BuildGetMailboxStatisticsScript(identity);

        onLog?.Invoke("Verbose", "Fetching mailbox statistics...");

        var result = await Engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        return result.Success && result.Output.FirstOrDefault()?.BaseObject is Hashtable hash
            ? ExoMailboxMapper.ToMailboxStatistics(hash)
            : null;
    }

    private async Task<List<InboxRuleDto>?> GetInboxRulesAsync(
        string identity,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var script = ExoMailboxScriptFactory.BuildGetInboxRulesScript(identity);

        onLog?.Invoke("Verbose", "Fetching inbox rules...");

        var result = await Engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success || !result.Output.Any())
        {
            return null;
        }

        var rules = new List<InboxRuleDto>();

        foreach (var output in result.Output)
        {
            if (output.BaseObject is object[] array)
            {
                foreach (var ruleHash in array.OfType<Hashtable>())
                {
                    rules.Add(ExoMailboxMapper.ToInboxRule(ruleHash));
                }

                continue;
            }

            if (output.BaseObject is Hashtable hash)
            {
                rules.Add(ExoMailboxMapper.ToInboxRule(hash));
            }
        }

        return rules;
    }

    private async Task<AutoReplyConfigurationDto?> GetAutoReplyConfigurationAsync(
        string identity,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var script = ExoMailboxScriptFactory.BuildGetAutoReplyConfigurationScript(identity);

        onLog?.Invoke("Verbose", "Fetching auto-reply configuration...");

        var result = await Engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        return result.Success && result.Output.FirstOrDefault()?.BaseObject is Hashtable hash
            ? ExoMailboxMapper.ToAutoReplyConfiguration(hash)
            : null;
    }
}
