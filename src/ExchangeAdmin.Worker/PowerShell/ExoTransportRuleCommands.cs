using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

internal sealed class ExoTransportRuleCommands : ExoCommandModuleBase
{
    public ExoTransportRuleCommands(PowerShellEngine engine)
        : base(engine)
    {
    }

    public async Task<GetTransportRulesResponse> GetTransportRulesAsync(CancellationToken cancellationToken = default)
    {
        var script = @"
$rules = Get-TransportRule -ErrorAction Stop | Sort-Object Priority
foreach ($r in $rules) {
    [PSCustomObject]@{
        Identity = $r.Identity.ToString()
        Name = $r.Name
        Priority = $r.Priority
        State = if ($r.State) { $r.State.ToString() } else { if ($r.Enabled) { 'Enabled' } else { 'Disabled' } }
        Mode = if ($r.Mode) { $r.Mode.ToString() } else { '' }
        Description = if ($r.Description) { $r.Description } else { '' }
        From = @($r.From)
        SentTo = @($r.SentTo)
        SenderDomainIs = @($r.SenderDomainIs)
        RecipientDomainIs = @($r.RecipientDomainIs)
        SentToMemberOf = @($r.SentToMemberOf)
        SubjectContainsWords = @($r.SubjectContainsWords)
        ExceptIfFrom = @($r.ExceptIfFrom)
        ExceptIfSentTo = @($r.ExceptIfSentTo)
        ExceptIfSenderDomainIs = @($r.ExceptIfSenderDomainIs)
        ExceptIfRecipientDomainIs = @($r.ExceptIfRecipientDomainIs)
        ExceptIfSubjectContainsWords = @($r.ExceptIfSubjectContainsWords)
        PrependSubject = if ($r.PrependSubject) { $r.PrependSubject } else { '' }
        RedirectMessageTo = @($r.RedirectMessageTo)
        BlindCopyTo = @($r.BlindCopyTo)
        AddToRecipients = @($r.AddToRecipients)
        StopRuleProcessing = [bool]$r.StopRuleProcessing
        DeleteMessage = [bool]$r.DeleteMessage
    }
}";

        var results = await RunScriptAsync(script, cancellationToken);
        var rules = new List<TransportRuleDto>();

        foreach (var obj in results)
        {
            rules.Add(new TransportRuleDto
            {
                Identity = GetString(obj, "Identity"),
                Name = GetString(obj, "Name"),
                Priority = obj.Properties["Priority"]?.Value == null ? null : Convert.ToInt32(obj.Properties["Priority"]?.Value),
                State = GetString(obj, "State"),
                Mode = GetString(obj, "Mode"),
                Description = GetString(obj, "Description"),
                From = ConvertToStringList(obj.Properties["From"]?.Value),
                SentTo = ConvertToStringList(obj.Properties["SentTo"]?.Value),
                SenderDomainIs = ConvertToStringList(obj.Properties["SenderDomainIs"]?.Value),
                RecipientDomainIs = ConvertToStringList(obj.Properties["RecipientDomainIs"]?.Value),
                SentToMemberOf = ConvertToStringList(obj.Properties["SentToMemberOf"]?.Value),
                SubjectContainsWords = ConvertToStringList(obj.Properties["SubjectContainsWords"]?.Value),
                ExceptIfFrom = ConvertToStringList(obj.Properties["ExceptIfFrom"]?.Value),
                ExceptIfSentTo = ConvertToStringList(obj.Properties["ExceptIfSentTo"]?.Value),
                ExceptIfSenderDomainIs = ConvertToStringList(obj.Properties["ExceptIfSenderDomainIs"]?.Value),
                ExceptIfRecipientDomainIs = ConvertToStringList(obj.Properties["ExceptIfRecipientDomainIs"]?.Value),
                ExceptIfSubjectContainsWords = ConvertToStringList(obj.Properties["ExceptIfSubjectContainsWords"]?.Value),
                PrependSubject = GetString(obj, "PrependSubject"),
                RedirectMessageTo = ConvertToStringList(obj.Properties["RedirectMessageTo"]?.Value),
                BlindCopyTo = ConvertToStringList(obj.Properties["BlindCopyTo"]?.Value),
                AddToRecipients = ConvertToStringList(obj.Properties["AddToRecipients"]?.Value),
                StopRuleProcessing = GetBool(obj, "StopRuleProcessing"),
                DeleteMessage = GetBool(obj, "DeleteMessage")
            });
        }

        return new GetTransportRulesResponse { Rules = rules };
    }

    public async Task SetTransportRuleStateAsync(SetTransportRuleStateRequest request, CancellationToken cancellationToken = default)
    {
        var escapedIdentity = EscapePs(request.Identity);
        var cmd = request.Enabled ? "Enable-TransportRule" : "Disable-TransportRule";
        var script = $@"
{cmd} -Identity '{escapedIdentity}' -Confirm:$false -ErrorAction Stop
Write-Output 'OK'";

        await RunScriptAsync(script, cancellationToken);
    }

    public async Task UpsertTransportRuleAsync(UpsertTransportRuleRequest request, CancellationToken cancellationToken = default)
    {
        var script = TransportRuleCommandBuilder.BuildUpsertTransportRuleScript(request);
        await RunScriptAsync(script, cancellationToken);
    }

    public async Task RemoveTransportRuleAsync(RemoveTransportRuleRequest request, CancellationToken cancellationToken = default)
    {
        var identity = EscapePs(request.Identity);
        var script = $@"
Remove-TransportRule -Identity '{identity}' -Confirm:$false -ErrorAction Stop
Write-Output 'OK'";

        await RunScriptAsync(script, cancellationToken);
    }

    public async Task<TestTransportRuleResponse> TestTransportRuleAsync(TestTransportRuleRequest request, CancellationToken cancellationToken = default)
    {
        var script = TransportRuleCommandBuilder.BuildTestTransportRuleScript(request);
        var results = await RunScriptAllowErrorsAsync(script, cancellationToken: cancellationToken);
        var names = results
            .Select(result => GetString(result, "Name"))
            .Where(static value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        return new TestTransportRuleResponse
        {
            MatchedRuleNames = names
        };
    }
}
