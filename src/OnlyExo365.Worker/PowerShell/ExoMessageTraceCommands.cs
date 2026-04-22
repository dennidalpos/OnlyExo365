using OnlyExo365.Contracts.Dtos;
using System.Management.Automation;

namespace OnlyExo365.Worker.PowerShell;

internal sealed class ExoMessageTraceCommands : ExoCommandModuleBase
{
    private const int MessageTraceV2BatchSize = 5000;
    private const int LegacyMessageTraceBatchSize = 5000;

    public ExoMessageTraceCommands(PowerShellEngine engine)
        : base(engine)
    {
    }

    public async Task<GetMessageTraceResponse> GetMessageTraceAsync(
        GetMessageTraceRequest request,
        CancellationToken cancellationToken = default)
    {
        var script = BuildGetMessageTraceScript(request);

        var result = await Engine.ExecuteAsync(script, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException(result.ErrorMessage ?? "Message trace query failed");
        }

        var warningDetails = ParseStructuredWarnings(result.Warning);
        var messages = new List<MessageTraceItemDto>();
        string? sourceCmdlet = null;

        foreach (var obj in result.Output)
        {
            if (string.Equals(GetNullableString(obj, "RecordType"), "Summary", StringComparison.Ordinal))
            {
                continue;
            }

            sourceCmdlet ??= GetNullableString(obj, "SourceCmdlet");
            messages.Add(new MessageTraceItemDto
            {
                MessageId = GetString(obj, "MessageId"),
                MessageTraceId = GetString(obj, "MessageTraceId"),
                SenderAddress = GetString(obj, "SenderAddress"),
                RecipientAddress = GetString(obj, "RecipientAddress"),
                Subject = GetString(obj, "Subject"),
                Status = GetString(obj, "Status"),
                Received = GetNullableDateTime(obj, "Received"),
                Size = GetNullableLong(obj, "Size")
            });
        }

        var totalCount = GetSummaryNullableInt(result.Output.LastOrDefault(), "TotalCount") ?? messages.Count;
        var isTotalCountExact = GetSummaryNullableBool(result.Output.LastOrDefault(), "IsTotalCountExact") ?? true;
        var pageStartIndex = Math.Max(0, (request.Page - 1) * request.PageSize);
        var hasMore = totalCount > pageStartIndex + messages.Count;

        return new GetMessageTraceResponse
        {
            Messages = messages,
            TotalCount = totalCount,
            HasMore = hasMore,
            Page = request.Page,
            PageSize = request.PageSize,
            Warnings = ExtractWarningMessages(warningDetails),
            WarningDetails = warningDetails,
            HasPartialData = warningDetails.Any(static warning => warning.IsPartialData),
            IsTotalCountExact = isTotalCountExact,
            SourceCmdlet = sourceCmdlet
        };
    }

    internal static string BuildGetMessageTraceScript(GetMessageTraceRequest request)
    {
        var escapedSenderAddress = EscapePs(request.SenderAddress);
        var escapedRecipientAddress = EscapePs(request.RecipientAddress);

        return $@"
$pageSize = {request.PageSize}
$page = {request.Page}
$requestedStartIndex = (($page - 1) * $pageSize)
$requestedEndIndex = $requestedStartIndex + $pageSize
$senderAddress = '{escapedSenderAddress}'
$recipientAddress = '{escapedRecipientAddress}'
$resultItems = New-Object 'System.Collections.Generic.List[object]'
$totalCount = 0
$sourceCmdlet = $null
$isTotalCountExact = $true

function Add-TraceItem {{
    param(
        [Parameter(Mandatory = $true)]
        [object]$Trace,
        [Parameter(Mandatory = $true)]
        [string]$CmdletName
    )

    if ($totalCount -ge $requestedStartIndex -and $totalCount -lt $requestedEndIndex) {{
        $resultItems.Add([PSCustomObject]@{{
            RecordType = 'Item'
            SourceCmdlet = $CmdletName
            MessageId = $Trace.MessageId
            MessageTraceId = if ($null -ne $Trace.MessageTraceId) {{ $Trace.MessageTraceId.ToString() }} else {{ '' }}
            SenderAddress = $Trace.SenderAddress
            RecipientAddress = $Trace.RecipientAddress
            Subject = $Trace.Subject
            Status = $Trace.Status
            Received = if ($Trace.Received) {{ $Trace.Received.ToString('o') }} else {{ $null }}
            Size = $Trace.Size
        }}) | Out-Null
    }}

    $script:totalCount++
}}

function New-SummaryRecord {{
    [PSCustomObject]@{{
        RecordType = 'Summary'
        SourceCmdlet = $sourceCmdlet
        TotalCount = $totalCount
        IsTotalCountExact = $isTotalCountExact
    }}
}}

$params = @{{
    StartDate = [DateTime]::Parse('{request.StartDate:o}')
    EndDate = [DateTime]::Parse('{request.EndDate:o}')
}}

if ($senderAddress -ne '') {{ $params['SenderAddress'] = $senderAddress }}
if ($recipientAddress -ne '') {{ $params['RecipientAddress'] = $recipientAddress }}

if (Get-Command Get-MessageTraceV2 -ErrorAction SilentlyContinue) {{
    $sourceCmdlet = 'Get-MessageTraceV2'
    $cursorEndDate = $params['EndDate']
    $cursorRecipientAddress = $null

    while ($true) {{
        $batchParams = @{{
            StartDate = $params['StartDate']
            EndDate = $cursorEndDate
            ResultSize = {MessageTraceV2BatchSize}
        }}

        if ($params.ContainsKey('SenderAddress')) {{ $batchParams['SenderAddress'] = $params['SenderAddress'] }}
        if ($params.ContainsKey('RecipientAddress')) {{ $batchParams['RecipientAddress'] = $params['RecipientAddress'] }}
        if (-not [string]::IsNullOrWhiteSpace($cursorRecipientAddress)) {{
            $batchParams['StartingRecipientAddress'] = $cursorRecipientAddress
        }}

        $batch = @(Get-MessageTraceV2 @batchParams -ErrorAction Stop)
        if ($batch.Count -eq 0) {{
            break
        }}

        foreach ($trace in $batch) {{
            Add-TraceItem -Trace $trace -CmdletName $sourceCmdlet
        }}

        if ($batch.Count -lt {MessageTraceV2BatchSize}) {{
            break
        }}

        $lastTrace = $batch[-1]
        if ($null -eq $lastTrace.Received) {{
            throw 'Get-MessageTraceV2 returned a page without Received. Unable to calculate the next page deterministically.'
        }}

        $nextEndDate = [DateTime]$lastTrace.Received
        $nextRecipientAddress = [string]($lastTrace.RecipientAddress ?? '')
        if ($cursorEndDate -eq $nextEndDate -and [string]::Equals($cursorRecipientAddress, $nextRecipientAddress, [System.StringComparison]::Ordinal)) {{
            throw 'Get-MessageTraceV2 paging cursor did not advance. Narrow the date range before retrying the query.'
        }}

        $cursorEndDate = $nextEndDate
        $cursorRecipientAddress = $nextRecipientAddress
    }}
}} elseif (Get-Command Get-MessageTrace -ErrorAction SilentlyContinue) {{
    $sourceCmdlet = 'Get-MessageTrace'
    $warningPayload = @{{
        Code = 'LegacyMessageTraceCmdlet'
        Scope = 'MessageTrace'
        Message = 'Get-MessageTraceV2 is not available. Falling back to legacy Get-MessageTrace.'
        IsPartialData = $true
    }} | ConvertTo-Json -Compress -Depth 3
    Write-Warning '{StructuredWarningPrefix}' + $warningPayload

    $legacyPage = 1
    while ($true) {{
        $batch = @(Get-MessageTrace @params -Page $legacyPage -PageSize {LegacyMessageTraceBatchSize} -ErrorAction Stop)
        if ($batch.Count -eq 0) {{
            break
        }}

        foreach ($trace in $batch) {{
            Add-TraceItem -Trace $trace -CmdletName $sourceCmdlet
        }}

        if ($batch.Count -lt {LegacyMessageTraceBatchSize}) {{
            break
        }}

        $legacyPage++
    }}
}} else {{
    throw 'Neither Get-MessageTraceV2 nor Get-MessageTrace is available. Install/upgrade ExchangeOnlineManagement and reconnect.'
}}

$resultItems.Add((New-SummaryRecord)) | Out-Null
$resultItems";
    }

    public async Task<GetMessageTraceDetailsResponse> GetMessageTraceDetailsAsync(
        GetMessageTraceDetailsRequest request,
        CancellationToken cancellationToken = default)
    {
        var escapedMessageTraceId = EscapePs(request.MessageTraceId);
        var escapedRecipient = EscapePs(request.RecipientAddress);

        var script = $@"
$messageTraceId = '{escapedMessageTraceId}'
$recipientAddress = '{escapedRecipient}'

if (Get-Command Get-MessageTraceDetailV2 -ErrorAction SilentlyContinue) {{
    $sourceCmdlet = 'Get-MessageTraceDetailV2'
    $details = Get-MessageTraceDetailV2 -MessageTraceId $messageTraceId -RecipientAddress $recipientAddress -ErrorAction Stop
}} elseif (Get-Command Get-MessageTraceDetail -ErrorAction SilentlyContinue) {{
    $sourceCmdlet = 'Get-MessageTraceDetail'
    $warningPayload = @{{
        Code = 'LegacyMessageTraceDetailCmdlet'
        Scope = 'MessageTrace.Details'
        Message = 'Get-MessageTraceDetailV2 is not available. Falling back to legacy Get-MessageTraceDetail.'
        IsPartialData = $true
    }} | ConvertTo-Json -Compress -Depth 3
    Write-Warning '{StructuredWarningPrefix}' + $warningPayload
    $details = Get-MessageTraceDetail -MessageTraceId $messageTraceId -RecipientAddress $recipientAddress -ErrorAction Stop
}} else {{
    throw 'Neither Get-MessageTraceDetailV2 nor Get-MessageTraceDetail is available.'
}}

foreach ($d in $details) {{
    [PSCustomObject]@{{
        SourceCmdlet = $sourceCmdlet
        Date = if ($d.Date) {{ $d.Date.ToString('o') }} else {{ $null }}
        Event = $d.Event
        Action = $d.Action
        Detail = $d.Detail
        Data = $d.Data
    }}
}}";

        var result = await Engine.ExecuteAsync(script, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException(result.ErrorMessage ?? "Message trace detail query failed");
        }

        var warningDetails = ParseStructuredWarnings(result.Warning);
        var events = new List<MessageTraceDetailEventDto>();
        string? sourceCmdlet = null;

        foreach (var obj in result.Output)
        {
            sourceCmdlet ??= GetNullableString(obj, "SourceCmdlet");
            events.Add(new MessageTraceDetailEventDto
            {
                Date = GetNullableDateTime(obj, "Date"),
                Event = GetString(obj, "Event"),
                Action = GetString(obj, "Action"),
                Detail = GetString(obj, "Detail"),
                Data = GetString(obj, "Data")
            });
        }

        return new GetMessageTraceDetailsResponse
        {
            MessageTraceId = request.MessageTraceId,
            Events = events,
            Warnings = ExtractWarningMessages(warningDetails),
            WarningDetails = warningDetails,
            HasPartialData = warningDetails.Any(static warning => warning.IsPartialData),
            SourceCmdlet = sourceCmdlet
        };
    }

    private static bool? GetSummaryNullableBool(PSObject? obj, string propertyName)
    {
        var value = GetRawPropertyValue(obj, propertyName);
        return value switch
        {
            null => null,
            bool boolValue => boolValue,
            _ when bool.TryParse(value.ToString(), out var parsed) => parsed,
            _ => null
        };
    }

    private static int? GetSummaryNullableInt(PSObject? obj, string propertyName)
    {
        var value = GetRawPropertyValue(obj, propertyName);
        return value switch
        {
            null => null,
            int intValue => intValue,
            long longValue when longValue is <= int.MaxValue and >= int.MinValue => (int)longValue,
            _ when int.TryParse(value.ToString(), out var parsed) => parsed,
            _ => null
        };
    }

    private static object? GetRawPropertyValue(PSObject? obj, string propertyName)
        => obj?.Properties[propertyName]?.Value;
}

