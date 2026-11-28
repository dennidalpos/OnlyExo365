using System.Text;
using OnlyExo365.Contracts.Dtos;

namespace OnlyExo365.Worker.PowerShell;

public static class TransportRuleCommandBuilder
{
    public static string BuildUpsertTransportRuleScript(UpsertTransportRuleRequest request)
    {
        var escapedIdentity = EscapePs(request.Identity);
        var escapedName = EscapePs(request.Name);
        var escapedPrepend = EscapePs(request.PrependSubject);
        var escapedMode = EscapePs(string.IsNullOrWhiteSpace(request.Mode) ? "Enforce" : request.Mode);

        var script = new StringBuilder();
        script.AppendLine($"$from = {ToPsArrayLiteral(request.From)}");
        script.AppendLine($"$sentTo = {ToPsArrayLiteral(request.SentTo)}");
        script.AppendLine($"$senderDomainIs = {ToPsArrayLiteral(request.SenderDomainIs)}");
        script.AppendLine($"$recipientDomainIs = {ToPsArrayLiteral(request.RecipientDomainIs)}");
        script.AppendLine($"$sentToMemberOf = {ToPsArrayLiteral(request.SentToMemberOf)}");
        script.AppendLine($"$subjectContains = {ToPsArrayLiteral(request.SubjectContainsWords)}");
        script.AppendLine($"$exceptIfFrom = {ToPsArrayLiteral(request.ExceptIfFrom)}");
        script.AppendLine($"$exceptIfSentTo = {ToPsArrayLiteral(request.ExceptIfSentTo)}");
        script.AppendLine($"$exceptIfSenderDomainIs = {ToPsArrayLiteral(request.ExceptIfSenderDomainIs)}");
        script.AppendLine($"$exceptIfRecipientDomainIs = {ToPsArrayLiteral(request.ExceptIfRecipientDomainIs)}");
        script.AppendLine($"$exceptIfSubjectContains = {ToPsArrayLiteral(request.ExceptIfSubjectContainsWords)}");
        script.AppendLine($"$redirectMessageTo = {ToPsArrayLiteral(request.RedirectMessageTo)}");
        script.AppendLine($"$blindCopyTo = {ToPsArrayLiteral(request.BlindCopyTo)}");
        script.AppendLine($"$addToRecipients = {ToPsArrayLiteral(request.AddToRecipients)}");
        script.AppendLine("$params = @{");
        script.AppendLine($"    Name = '{escapedName}'");
        script.AppendLine("}");
        var isUpdate = !string.IsNullOrWhiteSpace(request.Identity);
        AddArrayParameter(script, "From", "from", request.From, isUpdate);
        AddArrayParameter(script, "SentTo", "sentTo", request.SentTo, isUpdate);
        AddArrayParameter(script, "SenderDomainIs", "senderDomainIs", request.SenderDomainIs, isUpdate);
        AddArrayParameter(script, "RecipientDomainIs", "recipientDomainIs", request.RecipientDomainIs, isUpdate);
        AddArrayParameter(script, "SentToMemberOf", "sentToMemberOf", request.SentToMemberOf, isUpdate);
        AddArrayParameter(script, "SubjectContainsWords", "subjectContains", request.SubjectContainsWords, isUpdate);
        AddArrayParameter(script, "ExceptIfFrom", "exceptIfFrom", request.ExceptIfFrom, isUpdate);
        AddArrayParameter(script, "ExceptIfSentTo", "exceptIfSentTo", request.ExceptIfSentTo, isUpdate);
        AddArrayParameter(script, "ExceptIfSenderDomainIs", "exceptIfSenderDomainIs", request.ExceptIfSenderDomainIs, isUpdate);
        AddArrayParameter(script, "ExceptIfRecipientDomainIs", "exceptIfRecipientDomainIs", request.ExceptIfRecipientDomainIs, isUpdate);
        AddArrayParameter(script, "ExceptIfSubjectContainsWords", "exceptIfSubjectContains", request.ExceptIfSubjectContainsWords, isUpdate);
        AddArrayParameter(script, "RedirectMessageTo", "redirectMessageTo", request.RedirectMessageTo, isUpdate);
        AddArrayParameter(script, "BlindCopyTo", "blindCopyTo", request.BlindCopyTo, isUpdate);
        AddArrayParameter(script, "AddToRecipients", "addToRecipients", request.AddToRecipients, isUpdate);
        if (request.PrependSubject != null && (isUpdate || !string.IsNullOrWhiteSpace(request.PrependSubject)))
        {
            var value = string.IsNullOrWhiteSpace(request.PrependSubject) ? "$null" : $"'{escapedPrepend}'";
            script.AppendLine($"$params['PrependSubject'] = {value}");
        }
        if (request.StopRuleProcessing.HasValue)
            script.AppendLine($"$params['StopRuleProcessing'] = {ToPsBoolLiteral(request.StopRuleProcessing.Value)}");
        if (request.DeleteMessage.HasValue)
            script.AppendLine($"$params['DeleteMessage'] = {ToPsBoolLiteral(request.DeleteMessage.Value)}");
        if (!isUpdate || !string.IsNullOrWhiteSpace(request.Mode))
            script.AppendLine($"$params['Mode'] = '{escapedMode}'");

        if (isUpdate)
        {
            script.AppendLine($"Set-TransportRule -Identity '{escapedIdentity}' @params -ErrorAction Stop");
            // Set-TransportRule has no Enabled parameter; use the dedicated state cmdlet.
            if (request.Enabled.HasValue)
            {
                var command = request.Enabled.Value ? "Enable-TransportRule" : "Disable-TransportRule";
                script.AppendLine($"{command} -Identity '{escapedIdentity}' -Confirm:$false -ErrorAction Stop");
            }
        }
        else
        {
            script.AppendLine($"$params['Enabled'] = {ToPsBoolLiteral(request.Enabled ?? true)}");
            script.AppendLine("New-TransportRule @params -ErrorAction Stop");
        }

        return script.ToString();
    }

    private static void AddArrayParameter(StringBuilder script, string parameter, string variable, List<string>? values, bool isUpdate)
    {
        if (values == null) return;
        script.AppendLine($"if (${variable}.Count -gt 0) {{ $params['{parameter}'] = ${variable} }}");
        if (isUpdate)
            script.AppendLine($"else {{ $params['{parameter}'] = $null }}");
    }

    public static string BuildTestTransportRuleScript(TestTransportRuleRequest request)
    {
        var sender = EscapePs(request.Sender);
        var recipient = EscapePs(request.Recipient);
        var subject = EscapePs(request.Subject);

        return $@"
# NOTE: Exchange Online does not provide a fully reliable generic simulation cmdlet for all predicates.
# This is a best-effort matcher for the subset exposed by the UI.
$sender = '{sender}'
$recipient = '{recipient}'
$subject = '{subject}'
$senderDomain = if ($sender -like '*@*') {{ ($sender.Split('@')[-1]).ToLowerInvariant() }} else {{ '' }}
$recipientDomain = if ($recipient -like '*@*') {{ ($recipient.Split('@')[-1]).ToLowerInvariant() }} else {{ '' }}
$rules = Get-TransportRule -ErrorAction Stop
foreach ($r in $rules) {{
    $matches = $true
    $limitations = [System.Collections.Generic.List[string]]::new()
    if ($r.State -ne 'Enabled') {{ $limitations.Add('Rule state is missing or is not Enabled.') }}
    if ($r.Mode -ne 'Enforce') {{ $limitations.Add('Rule is in audit mode or its mode is unknown.') }}
    if (@($r.SentToMemberOf).Count -gt 0 -and $r.SentToMemberOf) {{ $limitations.Add('Group membership was not resolved.') }}
    $supported = @('From','SentTo','SenderDomainIs','RecipientDomainIs','SubjectContainsWords','ExceptIfFrom','ExceptIfSentTo','ExceptIfSenderDomainIs','ExceptIfRecipientDomainIs','ExceptIfSubjectContainsWords')
    foreach ($kind in @('Conditions','Exceptions')) {{
        if ($null -eq $r.PSObject.Properties[$kind]) {{
            $limitations.Add(""Condition metadata is unavailable: $kind."")
        }} else {{
            foreach ($predicate in @($r.$kind)) {{
                if ($null -ne $predicate -and $supported -notcontains $predicate.Name) {{
                    $limitations.Add(""Unsupported predicate: $($predicate.Name)."")
                }}
            }}
        }}
    }}
    foreach ($identities in @('From','SentTo','ExceptIfFrom','ExceptIfSentTo')) {{
        foreach ($identity in @($r.$identities)) {{
            if ($identity -and $identity.ToString() -notlike '*@*') {{ $limitations.Add('Recipient or sender identity was not resolved to SMTP.') }}
        }}
    }}

    if ($r.From -and $r.From.Count -gt 0) {{
        $matches = $matches -and (@($r.From) -contains $sender)
    }}

    if ($r.SentTo -and $r.SentTo.Count -gt 0) {{
        $matches = $matches -and (@($r.SentTo) -contains $recipient)
    }}

    if ($r.SenderDomainIs -and $r.SenderDomainIs.Count -gt 0) {{
        $ruleSenderDomains = @($r.SenderDomainIs) | ForEach-Object {{ $_.ToString().ToLowerInvariant() }}
        $matches = $matches -and ($ruleSenderDomains -contains $senderDomain)
    }}

    if ($r.RecipientDomainIs -and $r.RecipientDomainIs.Count -gt 0) {{
        $ruleRecipientDomains = @($r.RecipientDomainIs) | ForEach-Object {{ $_.ToString().ToLowerInvariant() }}
        $matches = $matches -and ($ruleRecipientDomains -contains $recipientDomain)
    }}

    if ($r.SubjectContainsWords -and $r.SubjectContainsWords.Count -gt 0) {{
        $subjectWords = @($r.SubjectContainsWords)
        $subjectMatch = ($subjectWords | Where-Object {{ $subject.IndexOf($_.ToString(), [System.StringComparison]::OrdinalIgnoreCase) -ge 0 }} | Measure-Object).Count -gt 0
        $matches = $matches -and $subjectMatch
    }}

    if ($r.ExceptIfFrom -and $r.ExceptIfFrom.Count -gt 0) {{
        if (@($r.ExceptIfFrom) -contains $sender) {{ $matches = $false }}
    }}

    if ($r.ExceptIfSentTo -and $r.ExceptIfSentTo.Count -gt 0) {{
        if (@($r.ExceptIfSentTo) -contains $recipient) {{ $matches = $false }}
    }}

    if ($r.ExceptIfSenderDomainIs -and $r.ExceptIfSenderDomainIs.Count -gt 0) {{
        $ruleExceptSenderDomains = @($r.ExceptIfSenderDomainIs) | ForEach-Object {{ $_.ToString().ToLowerInvariant() }}
        if ($ruleExceptSenderDomains -contains $senderDomain) {{ $matches = $false }}
    }}

    if ($r.ExceptIfRecipientDomainIs -and $r.ExceptIfRecipientDomainIs.Count -gt 0) {{
        $ruleExceptRecipientDomains = @($r.ExceptIfRecipientDomainIs) | ForEach-Object {{ $_.ToString().ToLowerInvariant() }}
        if ($ruleExceptRecipientDomains -contains $recipientDomain) {{ $matches = $false }}
    }}

    if ($r.ExceptIfSubjectContainsWords -and $r.ExceptIfSubjectContainsWords.Count -gt 0) {{
        $exceptWords = @($r.ExceptIfSubjectContainsWords)
        if (($exceptWords | Where-Object {{ $subject.IndexOf($_.ToString(), [System.StringComparison]::OrdinalIgnoreCase) -ge 0 }} | Measure-Object).Count -gt 0) {{
            $matches = $false
        }}
    }}

    $status = if ($r.State -eq 'Disabled') {{ 'NoMatch' }} elseif ($limitations.Count -gt 0) {{ 'Indeterminate' }} elseif (-not $matches) {{ 'NoMatch' }} else {{ 'Match' }}
    [PSCustomObject]@{{ Name = $r.Name; Status = $status; Reasons = @($limitations.ToArray()) }}
}}";
    }

    private static string EscapePs(string? value)
        => (value ?? string.Empty).Replace("'", "''");

    private static string ToPsBoolLiteral(bool value)
        => value ? "$true" : "$false";

    private static string ToPsArrayLiteral(IEnumerable<string>? values)
    {
        if (values == null)
        {
            return "@()";
        }

        var normalized = values
            .Where(static value => !string.IsNullOrWhiteSpace(value))
            .Select(static value => $"'{EscapePs(value.Trim())}'")
            .ToArray();

        return normalized.Length == 0
            ? "@()"
            : "@(" + string.Join(", ", normalized) + ")";
    }
}

