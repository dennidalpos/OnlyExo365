using System.Collections;
using System.Management.Automation;
using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

internal sealed class ExoMigrationCommands : ExoCommandModuleBase
{
    public ExoMigrationCommands(PowerShellEngine engine)
        : base(engine)
    {
    }

    public async Task<GetMigrationBatchesResponse> GetMigrationBatchesAsync(
        GetMigrationBatchesRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var normalizedStatus = NormalizeMigrationBatchStatus(request.Status);
        var sortProperty = NormalizeMigrationBatchSortProperty(request.SortBy);

        var response = new GetMigrationBatchesResponse
        {
            Skip = request.Skip,
            PageSize = request.PageSize,
            SearchQuery = request.SearchQuery
        };

        if (!string.IsNullOrWhiteSpace(request.Status) && normalizedStatus == null)
        {
            onLog?.Invoke("Warning", $"Unsupported migration batch status ignored: {request.Status}");
        }

        if (!string.IsNullOrWhiteSpace(request.SortBy) &&
            !string.Equals(sortProperty, request.SortBy.Trim(), StringComparison.OrdinalIgnoreCase))
        {
            onLog?.Invoke("Warning", $"Unsupported SortBy ignored: {request.SortBy}");
        }

        var escapedSearch = EscapePs(request.SearchQuery);
        var escapedStatus = EscapePs(normalizedStatus);
        var sortDirection = request.SortDescending ? "-Descending" : string.Empty;

        var script = $@"
$items = Get-MigrationBatch -ResultSize Unlimited -ErrorAction Stop | ForEach-Object {{
    [PSCustomObject]@{{
        Identity = $_.Identity.ToString()
        Name = if ($_.Name) {{ $_.Name.ToString() }} else {{ $_.Identity.ToString() }}
        Status = if ($_.Status) {{ $_.Status.ToString() }} else {{ '' }}
        State = if ($_.State) {{ $_.State.ToString() }} else {{ $null }}
        BatchType = if ($_.BatchType) {{ $_.BatchType.ToString() }} else {{ $null }}
        SourceEndpoint = if ($_.SourceEndpoint) {{ $_.SourceEndpoint.ToString() }} else {{ $null }}
        TargetEndpoint = if ($_.TargetEndpoint) {{ $_.TargetEndpoint.ToString() }} else {{ $null }}
        TotalCount = if ($null -ne $_.TotalCount) {{ [int]$_.TotalCount }} else {{ $null }}
        ActiveCount = if ($null -ne $_.ActiveCount) {{ [int]$_.ActiveCount }} else {{ $null }}
        SyncedCount = if ($null -ne $_.SyncedCount) {{ [int]$_.SyncedCount }} else {{ $null }}
        FinalizedCount = if ($null -ne $_.FinalizedCount) {{ [int]$_.FinalizedCount }} else {{ $null }}
        FailedCount = if ($null -ne $_.FailedCount) {{ [int]$_.FailedCount }} else {{ $null }}
        StoppedCount = if ($null -ne $_.StoppedCount) {{ [int]$_.StoppedCount }} else {{ $null }}
        CreatedBy = if ($_.CreatedBy) {{ $_.CreatedBy.ToString() }} else {{ $null }}
        CreatedDateTime = if ($_.CreatedDateTime) {{ $_.CreatedDateTime }} else {{ $null }}
        StartDateTime = if ($_.StartDateTime) {{ $_.StartDateTime }} else {{ $null }}
        CompleteDateTime = if ($_.CompleteDateTime) {{ $_.CompleteDateTime }} else {{ $null }}
        LastSyncedDateTime = if ($_.LastSyncedDateTime) {{ $_.LastSyncedDateTime }} else {{ $null }}
    }}
}}

$searchQuery = '{escapedSearch}'
if (-not [string]::IsNullOrWhiteSpace($searchQuery)) {{
    $items = $items | Where-Object {{
        $_.Name -like ""*$searchQuery*"" -or
        $_.Identity -like ""*$searchQuery*"" -or
        $_.Status -like ""*$searchQuery*"" -or
        $_.BatchType -like ""*$searchQuery*"" -or
        $_.SourceEndpoint -like ""*$searchQuery*"" -or
        $_.TargetEndpoint -like ""*$searchQuery*""
    }}
}}

$status = '{escapedStatus}'
if (-not [string]::IsNullOrWhiteSpace($status)) {{
    $items = $items | Where-Object {{ $_.Status -eq $status }}
}}

$items = $items | Sort-Object {sortProperty} {sortDirection}
$totalCount = @($items).Count
$pagedItems = $items | Select-Object -Skip {request.Skip} -First {request.PageSize}

@{{
    TotalCount = $totalCount
    Batches = @($pagedItems)
}}";

        onLog?.Invoke("Verbose", $"Fetching migration batches (skip={request.Skip}, pageSize={request.PageSize}, status={normalizedStatus ?? "All"})...");

        var result = await Engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (result.Success && result.Output.Any() && result.Output.First().BaseObject is Hashtable hash)
        {
            response.TotalCount = Convert.ToInt32(hash["TotalCount"] ?? 0);

            if (hash["Batches"] is object[] batches)
            {
                foreach (var batchObject in batches)
                {
                    if (batchObject is not PSObject batchPs)
                    {
                        continue;
                    }

                    response.Batches.Add(ToMigrationBatchListItem(batchPs));
                }
            }

            response.HasMore = (request.Skip + response.Batches.Count) < response.TotalCount;
        }

        onLog?.Invoke("Information", $"Retrieved {response.Batches.Count} migration batches (total: {response.TotalCount})");
        return response;
    }

    public async Task<GetMigrationEndpointsResponse> GetMigrationEndpointsAsync(
        GetMigrationEndpointsRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var searchQuery = EscapePs(request.SearchQuery);
        var sortProperty = NormalizeMigrationEndpointSortProperty(request.SortBy);
        var sortDirection = request.SortDescending ? "-Descending" : string.Empty;

        var script = $@"
$items = Get-MigrationEndpoint -ErrorAction Stop | ForEach-Object {{
    [PSCustomObject]@{{
        Identity = $_.Identity.ToString()
        Name = $_.Identity.ToString()
        EndpointType = if ($_.EndpointType) {{ $_.EndpointType.ToString() }} else {{ '' }}
        RemoteServer = if ($_.RemoteServer) {{ $_.RemoteServer.ToString() }} else {{ $null }}
        RpcProxyServer = if ($_.RpcProxyServer) {{ $_.RpcProxyServer.ToString() }} else {{ $null }}
        ExchangeServer = if ($_.ExchangeServer) {{ $_.ExchangeServer.ToString() }} else {{ $null }}
        EmailAddress = if ($_.EmailAddress) {{ $_.EmailAddress.ToString() }} else {{ $null }}
        RemoteTenant = if ($_.RemoteTenant) {{ $_.RemoteTenant.ToString() }} else {{ $null }}
        Port = if ($null -ne $_.Port) {{ [int]$_.Port }} else {{ $null }}
        Security = if ($_.Security) {{ $_.Security.ToString() }} else {{ $null }}
        Authentication = if ($_.Authentication) {{ $_.Authentication.ToString() }} else {{ $null }}
        MaxConcurrentMigrations = if ($null -ne $_.MaxConcurrentMigrations) {{ [int]$_.MaxConcurrentMigrations }} else {{ $null }}
        MaxConcurrentIncrementalSyncs = if ($null -ne $_.MaxConcurrentIncrementalSyncs) {{ [int]$_.MaxConcurrentIncrementalSyncs }} else {{ $null }}
        SkipVerification = if ($null -ne $_.SkipVerification) {{ [bool]$_.SkipVerification }} else {{ $null }}
        AcceptUntrustedCertificates = if ($null -ne $_.AcceptUntrustedCertificates) {{ [bool]$_.AcceptUntrustedCertificates }} else {{ $null }}
        LastModifiedTime = if ($_.LastModifiedTime) {{ $_.LastModifiedTime }} else {{ $null }}
    }}
}}

$searchQuery = '{searchQuery}'
if (-not [string]::IsNullOrWhiteSpace($searchQuery)) {{
    $items = $items | Where-Object {{
        $_.Name -like ""*$searchQuery*"" -or
        $_.EndpointType -like ""*$searchQuery*"" -or
        $_.RemoteServer -like ""*$searchQuery*"" -or
        $_.ExchangeServer -like ""*$searchQuery*"" -or
        $_.RpcProxyServer -like ""*$searchQuery*""
    }}
}}

$items | Sort-Object {sortProperty} {sortDirection}";

        onLog?.Invoke("Verbose", "Fetching migration endpoints...");
        var results = await RunScriptAsync(script, cancellationToken: cancellationToken);

        var response = new GetMigrationEndpointsResponse();
        foreach (var endpoint in results)
        {
            response.Endpoints.Add(ToMigrationEndpoint(endpoint));
        }

        onLog?.Invoke("Information", $"Retrieved {response.Endpoints.Count} migration endpoints");
        return response;
    }

    public async Task<MigrationBatchDetailsDto> GetMigrationBatchDetailsAsync(
        GetMigrationBatchDetailsRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var identity = EscapePs(request.Identity);
        var script = $@"
$batch = Get-MigrationBatch -Identity '{identity}' -IncludeReport -ErrorAction Stop

[PSCustomObject]@{{
    Identity = $batch.Identity.ToString()
    Name = if ($batch.Name) {{ $batch.Name.ToString() }} else {{ $batch.Identity.ToString() }}
    Status = if ($batch.Status) {{ $batch.Status.ToString() }} else {{ '' }}
    State = if ($batch.State) {{ $batch.State.ToString() }} else {{ $null }}
    StatusDetail = if ($batch.StatusDetail) {{ $batch.StatusDetail.ToString() }} else {{ $null }}
    BatchType = if ($batch.BatchType) {{ $batch.BatchType.ToString() }} else {{ $null }}
    SourceEndpoint = if ($batch.SourceEndpoint) {{ $batch.SourceEndpoint.ToString() }} else {{ $null }}
    TargetEndpoint = if ($batch.TargetEndpoint) {{ $batch.TargetEndpoint.ToString() }} else {{ $null }}
    TotalCount = if ($null -ne $batch.TotalCount) {{ [int]$batch.TotalCount }} else {{ $null }}
    ActiveCount = if ($null -ne $batch.ActiveCount) {{ [int]$batch.ActiveCount }} else {{ $null }}
    SyncedCount = if ($null -ne $batch.SyncedCount) {{ [int]$batch.SyncedCount }} else {{ $null }}
    FinalizedCount = if ($null -ne $batch.FinalizedCount) {{ [int]$batch.FinalizedCount }} else {{ $null }}
    FailedCount = if ($null -ne $batch.FailedCount) {{ [int]$batch.FailedCount }} else {{ $null }}
    StoppedCount = if ($null -ne $batch.StoppedCount) {{ [int]$batch.StoppedCount }} else {{ $null }}
    CreatedBy = if ($batch.CreatedBy) {{ $batch.CreatedBy.ToString() }} else {{ $null }}
    CreatedDateTime = if ($batch.CreatedDateTime) {{ $batch.CreatedDateTime }} else {{ $null }}
    StartDateTime = if ($batch.StartDateTime) {{ $batch.StartDateTime }} else {{ $null }}
    CompleteDateTime = if ($batch.CompleteDateTime) {{ $batch.CompleteDateTime }} else {{ $null }}
    LastSyncedDateTime = if ($batch.LastSyncedDateTime) {{ $batch.LastSyncedDateTime }} else {{ $null }}
    NotificationEmails = @($batch.NotificationEmails | ForEach-Object {{ $_.ToString() }})
    AutoStart = if ($null -ne $batch.AutoStart) {{ [bool]$batch.AutoStart }} else {{ $null }}
    AutoComplete = if ($null -ne $batch.AutoComplete) {{ [bool]$batch.AutoComplete }} else {{ $null }}
    BadItemLimit = if ($null -ne $batch.BadItemLimit) {{ [int]$batch.BadItemLimit }} else {{ $null }}
    LargeItemLimit = if ($null -ne $batch.LargeItemLimit) {{ [int]$batch.LargeItemLimit }} else {{ $null }}
    Report = if ($batch.Report) {{ $batch.Report.ToString() }} else {{ $null }}
    StartAfter = if ($batch.StartAfter) {{ $batch.StartAfter }} else {{ $null }}
    CompleteAfter = if ($batch.CompleteAfter) {{ $batch.CompleteAfter }} else {{ $null }}
}}";

        onLog?.Invoke("Verbose", $"Fetching migration batch details for {request.Identity}...");

        var results = await RunScriptAsync(script, cancellationToken);
        if (results.Count == 0)
        {
            throw new InvalidOperationException($"Migration batch not found: {request.Identity}");
        }

        var obj = results[0];
        return new MigrationBatchDetailsDto
        {
            Identity = GetString(obj, "Identity"),
            Name = GetString(obj, "Name"),
            Status = GetString(obj, "Status"),
            State = GetNullableString(obj, "State"),
            StatusDetail = GetNullableString(obj, "StatusDetail"),
            BatchType = GetNullableString(obj, "BatchType"),
            SourceEndpoint = GetNullableString(obj, "SourceEndpoint"),
            TargetEndpoint = GetNullableString(obj, "TargetEndpoint"),
            TotalCount = GetNullableInt(obj, "TotalCount"),
            ActiveCount = GetNullableInt(obj, "ActiveCount"),
            SyncedCount = GetNullableInt(obj, "SyncedCount"),
            FinalizedCount = GetNullableInt(obj, "FinalizedCount"),
            FailedCount = GetNullableInt(obj, "FailedCount"),
            StoppedCount = GetNullableInt(obj, "StoppedCount"),
            CreatedBy = GetNullableString(obj, "CreatedBy"),
            CreatedDateTime = GetNullableDateTime(obj, "CreatedDateTime"),
            StartDateTime = GetNullableDateTime(obj, "StartDateTime"),
            CompleteDateTime = GetNullableDateTime(obj, "CompleteDateTime"),
            LastSyncedDateTime = GetNullableDateTime(obj, "LastSyncedDateTime"),
            NotificationEmails = ConvertToStringList(obj.Properties["NotificationEmails"]?.Value),
            AutoStart = GetNullableBool(obj, "AutoStart"),
            AutoComplete = GetNullableBool(obj, "AutoComplete"),
            BadItemLimit = GetNullableInt(obj, "BadItemLimit"),
            LargeItemLimit = GetNullableInt(obj, "LargeItemLimit"),
            Report = GetNullableString(obj, "Report"),
            StartAfter = GetNullableDateTime(obj, "StartAfter"),
            CompleteAfter = GetNullableDateTime(obj, "CompleteAfter")
        };
    }

    public async Task UpsertMigrationEndpointAsync(
        UpsertMigrationEndpointRequest request,
        CancellationToken cancellationToken = default)
    {
        var command = MigrationCommandBuilder.BuildUpsertMigrationEndpointCommand(request);
        await RunScriptAsync(command.Script, command.Parameters, cancellationToken);
    }

    public async Task<TestMigrationEndpointResponse> TestMigrationEndpointAsync(
        TestMigrationEndpointRequest request,
        CancellationToken cancellationToken = default)
    {
        var command = MigrationCommandBuilder.BuildTestMigrationEndpointCommand(request);
        var results = await RunScriptAsync(command.Script, command.Parameters, cancellationToken);
        if (results.Count == 0)
        {
            return new TestMigrationEndpointResponse
            {
                Summary = "Migration endpoint test completed."
            };
        }

        var result = results[0];
        return new TestMigrationEndpointResponse
        {
            Summary = GetString(result, "Summary"),
            Details = GetNullableString(result, "Details")
        };
    }

    public async Task<GetMigrationBatchPreflightResponse> GetMigrationBatchPreflightAsync(
        GetMigrationBatchPreflightRequest request,
        CancellationToken cancellationToken = default)
    {
        var batchType = NormalizeMigrationBatchCreationType(request.BatchType);
        var identity = EscapePs(request.EndpointIdentity);
        var batchName = EscapePs(request.Name);
        var targetDeliveryDomain = EscapePs(request.TargetDeliveryDomain);
        var requiredColumns = batchType == "IMAP"
            ? "EmailAddress, UserName, Password"
            : "EmailAddress";

        var script = $@"
param(
    [string]$CsvFilePath
)

$messages = [System.Collections.Generic.List[string]]::new()
$headers = [System.Collections.Generic.List[string]]::new()
$rowCount = 0
$isReady = $true
$endpointType = $null

try {{
    $endpoint = Get-MigrationEndpoint -Identity '{identity}' -ErrorAction Stop
    $endpointType = if ($endpoint.EndpointType) {{ $endpoint.EndpointType.ToString() }} else {{ $null }}
}}
catch {{
    $messages.Add('Migration endpoint was not found or is not readable.')
    $isReady = $false
}}

if (-not [string]::IsNullOrWhiteSpace('{batchName}')) {{
    try {{
        $existingBatch = Get-MigrationBatch -Identity '{batchName}' -ErrorAction Stop
        if ($null -ne $existingBatch) {{
            $messages.Add('A migration batch with the same name already exists.')
            $isReady = $false
        }}
    }}
    catch {{
    }}
}}

if (-not [System.IO.File]::Exists($CsvFilePath)) {{
    $messages.Add('CSV file not found.')
    $isReady = $false
}}
else {{
    $headerLine = Get-Content -Path $CsvFilePath -First 1 -ErrorAction Stop
    if (-not [string]::IsNullOrWhiteSpace($headerLine)) {{
        foreach ($header in ($headerLine -split ',')) {{
            $trimmed = $header.Trim()
            if (-not [string]::IsNullOrWhiteSpace($trimmed)) {{
                [void]$headers.Add($trimmed)
            }}
        }}
    }}

    $rows = @(Import-Csv -Path $CsvFilePath -ErrorAction Stop)
    $rowCount = $rows.Count

    if ($rowCount -eq 0) {{
        $messages.Add('CSV file is empty: no data rows found.')
        $isReady = $false
    }}
}}

$requiredHeaders = @({ToPsArrayLiteral(requiredColumns.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries))})
foreach ($requiredHeader in $requiredHeaders) {{
    if ($headers -notcontains $requiredHeader) {{
        $messages.Add(""Required CSV column missing: $requiredHeader"")
        $isReady = $false
    }}
}}

$batchType = '{batchType}'
switch ($batchType) {{
    'IMAP' {{
        if ($endpointType -and $endpointType -ne 'IMAP') {{
            $messages.Add('The IMAP batch requires an IMAP endpoint.')
            $isReady = $false
        }}
    }}
    'Offboarding' {{
        if ($endpointType -and $endpointType -ne 'ExchangeRemoteMove') {{
            $messages.Add('The Offboarding batch requires an ExchangeRemoteMove endpoint.')
            $isReady = $false
        }}
    }}
    default {{
        if ($endpointType -and $endpointType -notin @('ExchangeRemoteMove', 'ExchangeOutlookAnywhere')) {{
            $messages.Add('The Onboarding batch requires an ExchangeRemoteMove or ExchangeOutlookAnywhere endpoint.')
            $isReady = $false
        }}
        if ([string]::IsNullOrWhiteSpace('{targetDeliveryDomain}')) {{
            $messages.Add('Target delivery domain is not set: the batch will be created without this parameter.')
        }}
    }}
}}

if ($isReady -and $messages.Count -eq 0) {{
    $messages.Add('Preflight completed successfully.')
}}

[PSCustomObject]@{{
    IsReady = $isReady
    EndpointType = $endpointType
    CsvRowCount = $rowCount
    CsvHeaders = @($headers)
    Messages = @($messages)
}}";

        var results = await RunScriptAsync(
            script,
            new Dictionary<string, object> { ["CsvFilePath"] = request.CsvFilePath },
            cancellationToken);

        if (results.Count == 0)
        {
            return new GetMigrationBatchPreflightResponse
            {
                IsReady = false,
                Messages = ["Migration preflight returned no results."]
            };
        }

        var result = results[0];
        return new GetMigrationBatchPreflightResponse
        {
            IsReady = GetBool(result, "IsReady"),
            EndpointType = GetNullableString(result, "EndpointType"),
            CsvRowCount = GetNullableInt(result, "CsvRowCount") ?? 0,
            CsvHeaders = ConvertToStringList(result.Properties["CsvHeaders"]?.Value),
            Messages = ConvertToStringList(result.Properties["Messages"]?.Value)
        };
    }

    public async Task CreateMigrationBatchAsync(
        CreateMigrationBatchRequest request,
        CancellationToken cancellationToken = default)
    {
        var command = MigrationCommandBuilder.BuildCreateMigrationBatchCommand(request);
        await RunScriptAsync(command.Script, command.Parameters, cancellationToken);
    }

    public async Task StartMigrationBatchAsync(StartMigrationBatchRequest request, CancellationToken cancellationToken = default)
    {
        var identity = EscapePs(request.Identity);
        var script = $@"
Start-MigrationBatch -Identity '{identity}' -Confirm:$false -ErrorAction Stop
Write-Output 'OK'";

        await RunScriptAsync(script, cancellationToken);
    }

    public async Task CompleteMigrationBatchAsync(CompleteMigrationBatchRequest request, CancellationToken cancellationToken = default)
    {
        var identity = EscapePs(request.Identity);
        var script = $@"
Complete-MigrationBatch -Identity '{identity}' -Confirm:$false -ErrorAction Stop
Write-Output 'OK'";

        await RunScriptAsync(script, cancellationToken);
    }

    public async Task RemoveMigrationBatchAsync(RemoveMigrationBatchRequest request, CancellationToken cancellationToken = default)
    {
        var identity = EscapePs(request.Identity);
        var script = $@"
Remove-MigrationBatch -Identity '{identity}' -Confirm:$false -ErrorAction Stop
Write-Output 'OK'";

        await RunScriptAsync(script, cancellationToken);
    }

    private static MigrationBatchListItemDto ToMigrationBatchListItem(PSObject batchPs)
    {
        return new MigrationBatchListItemDto
        {
            Identity = GetString(batchPs, "Identity"),
            Name = GetString(batchPs, "Name"),
            Status = GetString(batchPs, "Status"),
            State = GetNullableString(batchPs, "State"),
            BatchType = GetNullableString(batchPs, "BatchType"),
            SourceEndpoint = GetNullableString(batchPs, "SourceEndpoint"),
            TargetEndpoint = GetNullableString(batchPs, "TargetEndpoint"),
            TotalCount = GetNullableInt(batchPs, "TotalCount"),
            ActiveCount = GetNullableInt(batchPs, "ActiveCount"),
            SyncedCount = GetNullableInt(batchPs, "SyncedCount"),
            FinalizedCount = GetNullableInt(batchPs, "FinalizedCount"),
            FailedCount = GetNullableInt(batchPs, "FailedCount"),
            StoppedCount = GetNullableInt(batchPs, "StoppedCount"),
            CreatedBy = GetNullableString(batchPs, "CreatedBy"),
            CreatedDateTime = GetNullableDateTime(batchPs, "CreatedDateTime"),
            StartDateTime = GetNullableDateTime(batchPs, "StartDateTime"),
            CompleteDateTime = GetNullableDateTime(batchPs, "CompleteDateTime"),
            LastSyncedDateTime = GetNullableDateTime(batchPs, "LastSyncedDateTime")
        };
    }

    private static MigrationEndpointDto ToMigrationEndpoint(PSObject endpoint)
    {
        return new MigrationEndpointDto
        {
            Identity = GetString(endpoint, "Identity"),
            Name = GetString(endpoint, "Name"),
            EndpointType = GetString(endpoint, "EndpointType"),
            RemoteServer = GetNullableString(endpoint, "RemoteServer"),
            RpcProxyServer = GetNullableString(endpoint, "RpcProxyServer"),
            ExchangeServer = GetNullableString(endpoint, "ExchangeServer"),
            EmailAddress = GetNullableString(endpoint, "EmailAddress"),
            RemoteTenant = GetNullableString(endpoint, "RemoteTenant"),
            Port = GetNullableInt(endpoint, "Port"),
            Security = GetNullableString(endpoint, "Security"),
            Authentication = GetNullableString(endpoint, "Authentication"),
            MaxConcurrentMigrations = GetNullableInt(endpoint, "MaxConcurrentMigrations"),
            MaxConcurrentIncrementalSyncs = GetNullableInt(endpoint, "MaxConcurrentIncrementalSyncs"),
            SkipVerification = GetNullableBool(endpoint, "SkipVerification"),
            AcceptUntrustedCertificates = GetNullableBool(endpoint, "AcceptUntrustedCertificates"),
            LastModifiedTime = GetNullableDateTime(endpoint, "LastModifiedTime")
        };
    }

    private static string? NormalizeMigrationBatchStatus(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        return value.Trim() switch
        {
            "Created" => "Created",
            "Starting" => "Starting",
            "Syncing" => "Syncing",
            "Synced" => "Synced",
            "Completing" => "Completing",
            "Completed" => "Completed",
            "CompletedWithErrors" => "CompletedWithErrors",
            "Failed" => "Failed",
            "Stopped" => "Stopped",
            "Stopping" => "Stopping",
            "Removing" => "Removing",
            _ => null
        };
    }

    private static string NormalizeMigrationBatchSortProperty(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "Name";
        }

        return value.Trim() switch
        {
            "Status" => "Status",
            "BatchType" => "BatchType",
            "CreatedDateTime" => "CreatedDateTime",
            "StartDateTime" => "StartDateTime",
            "CompleteDateTime" => "CompleteDateTime",
            "SyncedCount" => "SyncedCount",
            "FailedCount" => "FailedCount",
            "TotalCount" => "TotalCount",
            _ => "Name"
        };
    }

    private static string NormalizeMigrationEndpointSortProperty(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "Name";
        }

        return value.Trim() switch
        {
            "EndpointType" => "EndpointType",
            "RemoteServer" => "RemoteServer",
            "LastModifiedTime" => "LastModifiedTime",
            _ => "Name"
        };
    }

    private static string NormalizeMigrationBatchCreationType(string? value)
        => value?.Trim() switch
        {
            "Offboarding" => "Offboarding",
            "IMAP" => "IMAP",
            _ => "Onboarding"
        };
}
