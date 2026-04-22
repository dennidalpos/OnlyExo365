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
        script.AppendLine($"    Mode = '{escapedMode}'");
        script.AppendLine($"    Enabled = {ToPsBoolLiteral(request.Enabled)}");
        script.AppendLine("}");
        script.AppendLine("if ($from.Count -gt 0) { $params['From'] = $from }");
        script.AppendLine("if ($sentTo.Count -gt 0) { $params['SentTo'] = $sentTo }");
        script.AppendLine("if ($senderDomainIs.Count -gt 0) { $params['SenderDomainIs'] = $senderDomainIs }");
        script.AppendLine("if ($recipientDomainIs.Count -gt 0) { $params['RecipientDomainIs'] = $recipientDomainIs }");
        script.AppendLine("if ($sentToMemberOf.Count -gt 0) { $params['SentToMemberOf'] = $sentToMemberOf }");
        script.AppendLine("if ($subjectContains.Count -gt 0) { $params['SubjectContainsWords'] = $subjectContains }");
        script.AppendLine("if ($exceptIfFrom.Count -gt 0) { $params['ExceptIfFrom'] = $exceptIfFrom }");
        script.AppendLine("if ($exceptIfSentTo.Count -gt 0) { $params['ExceptIfSentTo'] = $exceptIfSentTo }");
        script.AppendLine("if ($exceptIfSenderDomainIs.Count -gt 0) { $params['ExceptIfSenderDomainIs'] = $exceptIfSenderDomainIs }");
        script.AppendLine("if ($exceptIfRecipientDomainIs.Count -gt 0) { $params['ExceptIfRecipientDomainIs'] = $exceptIfRecipientDomainIs }");
        script.AppendLine("if ($exceptIfSubjectContains.Count -gt 0) { $params['ExceptIfSubjectContainsWords'] = $exceptIfSubjectContains }");
        script.AppendLine("if ($redirectMessageTo.Count -gt 0) { $params['RedirectMessageTo'] = $redirectMessageTo }");
        script.AppendLine("if ($blindCopyTo.Count -gt 0) { $params['BlindCopyTo'] = $blindCopyTo }");
        script.AppendLine("if ($addToRecipients.Count -gt 0) { $params['AddToRecipients'] = $addToRecipients }");
        script.AppendLine($"if ('{escapedPrepend}' -ne '') {{ $params['PrependSubject'] = '{escapedPrepend}' }}");
        script.AppendLine($"if ({ToPsBoolLiteral(request.StopRuleProcessing)}) {{ $params['StopRuleProcessing'] = $true }}");
        script.AppendLine($"if ({ToPsBoolLiteral(request.DeleteMessage)}) {{ $params['DeleteMessage'] = $true }}");
        script.AppendLine();
        script.AppendLine($"if ('{escapedIdentity}' -ne '') {{");
        script.AppendLine($"    Set-TransportRule -Identity '{escapedIdentity}' @params -ErrorAction Stop");
        script.AppendLine("} else {");
        script.AppendLine("    New-TransportRule @params -ErrorAction Stop");
        script.AppendLine("}");

        return script.ToString();
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
        $subjectMatch = ($subjectWords | Where-Object {{ $subject -like ""*$_*"" }} | Measure-Object).Count -gt 0
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
        if (($exceptWords | Where-Object {{ $subject -like ""*$_*"" }} | Measure-Object).Count -gt 0) {{
            $matches = $false
        }}
    }}

    if ($matches) {{
        [PSCustomObject]@{{ Name = $r.Name }}
    }}
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

