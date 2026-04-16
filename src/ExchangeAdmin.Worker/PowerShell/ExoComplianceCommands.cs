using System.Collections;
using System.Management.Automation;
using System.Text;
using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

internal sealed class ExoComplianceCommands : ExoCommandModuleBase
{
    private readonly ComplianceSearchOnlyRunner _searchOnlyRunner;

    public ExoComplianceCommands(PowerShellEngine engine)
        : base(engine)
    {
        _searchOnlyRunner = new ComplianceSearchOnlyRunner(engine.ExchangeConfiguration);
    }

    public async Task<GetComplianceWorkspaceResponse> GetComplianceWorkspaceAsync(
        GetComplianceWorkspaceRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        await EnsureComplianceCmdletAsync("Get-ComplianceSearch", onLog, cancellationToken);

        var result = await Engine.ExecuteAsync(
            BuildGetComplianceWorkspaceScript(request.MaxActions),
            onVerbose: onLog,
            cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException(result.ErrorMessage ?? "Unable to load compliance workspace.");
        }

        return MapWorkspaceResponse(result.Output, result.Warning);
    }

    internal static GetComplianceWorkspaceResponse MapWorkspaceResponse(
        IEnumerable<PSObject> results,
        IEnumerable<string>? warnings)
    {
        var warningDetails = ParseStructuredWarnings(warnings);
        var unsupportedHoldWarning = warningDetails.FirstOrDefault(static warning =>
            string.Equals(warning.Code, "CaseHoldPolicyUnavailable", StringComparison.Ordinal));
        var response = new GetComplianceWorkspaceResponse
        {
            WarningDetails = warningDetails,
            Warnings = ExtractWarningMessages(warningDetails),
            HasPartialData = warningDetails.Any(static warning => warning.IsPartialData),
            IsHoldListingUnsupported = unsupportedHoldWarning != null,
            HoldListingStatusMessage = unsupportedHoldWarning?.Message
        };

        var workspace = results.FirstOrDefault();
        if (workspace == null)
        {
            return response;
        }

        foreach (var obj in EnumeratePsObjects(GetPropertyValue(workspace, "Searches")))
        {
            response.Searches.Add(new ComplianceSearchDto
            {
                Name = GetString(obj, "Name"),
                CaseName = GetNullableString(obj, "CaseName"),
                Status = GetNullableString(obj, "Status"),
                CreatedBy = GetNullableString(obj, "CreatedBy"),
                CreatedTime = GetNullableDateTime(obj, "CreatedTime"),
                LastModifiedTime = GetNullableDateTime(obj, "LastModifiedTime"),
                Items = GetNullableString(obj, "Items"),
                Size = GetNullableString(obj, "Size"),
                ExchangeLocations = ConvertToStringList(GetPropertyValue(obj, "ExchangeLocations")),
                ContentMatchQuery = GetNullableString(obj, "ContentMatchQuery")
            });
        }

        foreach (var obj in EnumeratePsObjects(GetPropertyValue(workspace, "Cases")))
        {
            response.Cases.Add(new ComplianceCaseDto
            {
                Name = GetString(obj, "Name"),
                Status = GetNullableString(obj, "Status"),
                CaseType = GetNullableString(obj, "CaseType"),
                CreatedTime = GetNullableDateTime(obj, "CreatedTime"),
                Description = GetNullableString(obj, "Description")
            });
        }

        foreach (var obj in EnumeratePsObjects(GetPropertyValue(workspace, "Actions")))
        {
            response.Actions.Add(MapComplianceAction(obj));
        }

        return response;
    }

    public async Task<SearchUnifiedAuditLogResponse> SearchUnifiedAuditLogAsync(
        SearchUnifiedAuditLogRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var result = await _searchOnlyRunner.ExecuteAsync(
            BuildSearchUnifiedAuditLogScript(request),
            "Search-UnifiedAuditLog",
            onLog,
            cancellationToken).ConfigureAwait(false);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException(result.ErrorMessage ?? "Unable to search unified audit log.");
        }

        return MapSearchUnifiedAuditLogResponse(result.Output, request.MaxResults);
    }

    public async Task CreateComplianceSearchAsync(
        CreateComplianceSearchRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(request.Name))
        {
            throw new InvalidOperationException("Name is required.");
        }

        if (request.ExchangeLocations.Count == 0)
        {
            throw new InvalidOperationException("At least one Exchange location is required.");
        }

        await EnsureComplianceCmdletAsync("New-ComplianceSearch", onLog, cancellationToken);
        await RunScriptAsync(ComplianceCommandBuilder.BuildCreateComplianceSearchScript(request), cancellationToken);
    }

    public async Task StartComplianceSearchAsync(
        StartComplianceSearchRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(request.Name))
        {
            throw new InvalidOperationException("Name is required.");
        }

        await EnsureComplianceCmdletAsync("Start-ComplianceSearch", onLog, cancellationToken);
        await RunScriptAsync(ComplianceCommandBuilder.BuildStartComplianceSearchScript(request.Name), cancellationToken);
    }

    public async Task RemoveComplianceSearchAsync(
        RemoveComplianceSearchRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(request.Name))
        {
            throw new InvalidOperationException("Name is required.");
        }

        await EnsureComplianceCmdletAsync("Remove-ComplianceSearch", onLog, cancellationToken);
        await RunScriptAsync(ComplianceCommandBuilder.BuildRemoveComplianceSearchScript(request.Name), cancellationToken);
    }

    public async Task<InvokeComplianceActionResponse> InvokeComplianceActionAsync(
        InvokeComplianceActionRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(request.SearchName))
        {
            throw new InvalidOperationException("SearchName is required.");
        }

        var actionType = NormalizeActionType(request.ActionType);
        switch (actionType)
        {
            case "Purge":
                await EnsureComplianceCmdletAsync("New-ComplianceSearchAction", onLog, cancellationToken);
                break;
            case "Hold":
                await EnsureComplianceCmdletAsync("New-CaseHoldPolicy", onLog, cancellationToken);
                await EnsureComplianceCmdletAsync("New-CaseHoldRule", onLog, cancellationToken);
                break;
            default:
                throw new InvalidOperationException($"Unsupported compliance action type: {request.ActionType}");
        }

        var selectedSearch = await GetComplianceSearchAsync(request.SearchName, cancellationToken)
            ?? throw new InvalidOperationException($"Compliance search not found: {request.SearchName}");

        var script = ComplianceCommandBuilder.BuildInvokeComplianceActionScript(
            new InvokeComplianceActionRequest
            {
                SearchName = request.SearchName,
                ActionType = actionType,
                PurgeType = NormalizePurgeType(request.PurgeType),
                CaseName = request.CaseName,
                HoldName = request.HoldName
            },
            selectedSearch.ExchangeLocations,
            selectedSearch.ContentMatchQuery);

        var results = await RunScriptAsync(script, cancellationToken);
        var resultObject = results.FirstOrDefault();
        var action = resultObject == null
            ? CreateFallbackAction(request, selectedSearch)
            : actionType == "Hold"
                ? MapHoldAction(resultObject, request.SearchName, request.CaseName)
                : MapComplianceAction(resultObject);

        if (string.IsNullOrWhiteSpace(action.SearchName))
        {
            action.SearchName = request.SearchName;
        }

        if (string.IsNullOrWhiteSpace(action.CaseName))
        {
            action.CaseName = request.CaseName;
        }

        if (string.IsNullOrWhiteSpace(action.ActionType))
        {
            action.ActionType = actionType;
        }

        return new InvokeComplianceActionResponse
        {
            Action = action
        };
    }

    private async Task EnsureComplianceCmdletAsync(
        string cmdletName,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        if (!Engine.IsComplianceConnected)
        {
            onLog?.Invoke("Information", "Connecting to Security & Compliance PowerShell...");
            var connectResult = await Engine.ConnectComplianceAsync(onVerbose: onLog, cancellationToken: cancellationToken);
            if (!connectResult.Success)
            {
                throw new InvalidOperationException(connectResult.ErrorMessage ?? "Unable to connect to Security & Compliance PowerShell.");
            }
        }

        var validationResult = await ValidateComplianceCmdletAsync(cmdletName, cancellationToken);
        if (validationResult.Success)
        {
            return;
        }

        onLog?.Invoke("Warning", $"Compliance cmdlet {cmdletName} not available in the current session. Retrying Security & Compliance PowerShell connection...");
        Engine.ComplianceConnected = false;

        var retryConnectResult = await Engine.ConnectComplianceAsync(onVerbose: onLog, cancellationToken: cancellationToken);
        if (!retryConnectResult.Success)
        {
            throw new InvalidOperationException(retryConnectResult.ErrorMessage ?? "Unable to reconnect to Security & Compliance PowerShell.");
        }

        validationResult = await ValidateComplianceCmdletAsync(cmdletName, cancellationToken);
        if (!validationResult.Success)
        {
            throw new InvalidOperationException(validationResult.ErrorMessage ?? $"Cmdlet {cmdletName} is not available in the current Security & Compliance PowerShell session.");
        }
    }

    private Task<PowerShellResult> ValidateComplianceCmdletAsync(string cmdletName, CancellationToken cancellationToken)
    {
        var escapedCmdletName = EscapePs(cmdletName);
        var validationScript = $@"
if (-not (Get-Command -Name '{escapedCmdletName}' -ErrorAction SilentlyContinue)) {{
    throw 'Cmdlet {escapedCmdletName} is not available in the current Security & Compliance PowerShell session.'
}}";

        return Engine.ExecuteAsync(validationScript, cancellationToken: cancellationToken);
    }

    private async Task<ComplianceSearchDto?> GetComplianceSearchAsync(string searchName, CancellationToken cancellationToken)
    {
        var escapedSearchName = EscapePs(searchName);
        var script = $@"
function Get-StringList($value) {{
    $items = New-Object System.Collections.Generic.List[string]
    foreach ($item in @($value)) {{
        if ($null -eq $item) {{ continue }}
        $text = $item.ToString()
        if (-not [string]::IsNullOrWhiteSpace($text)) {{
            $items.Add($text)
        }}
    }}

    return @($items)
}}

$search = Get-ComplianceSearch -Identity '{escapedSearchName}' -ErrorAction Stop | Select-Object -First 1
[PSCustomObject]@{{
    Name = $search.Name
    CaseName = if ($search.Case) {{ $search.Case.ToString() }} else {{ $null }}
    Status = if ($search.Status) {{ $search.Status.ToString() }} else {{ $null }}
    ExchangeLocations = @(Get-StringList $search.ExchangeLocation)
    ContentMatchQuery = if ($search.ContentMatchQuery) {{ $search.ContentMatchQuery.ToString() }} else {{ $null }}
}}";

        var result = await RunScriptAsync(script, cancellationToken);
        var obj = result.FirstOrDefault();
        if (obj == null)
        {
            return null;
        }

        return new ComplianceSearchDto
        {
            Name = GetString(obj, "Name"),
            CaseName = GetNullableString(obj, "CaseName"),
            Status = GetNullableString(obj, "Status"),
            ExchangeLocations = ConvertToStringList(GetPropertyValue(obj, "ExchangeLocations")),
            ContentMatchQuery = GetNullableString(obj, "ContentMatchQuery")
        };
    }

    internal static string BuildGetComplianceWorkspaceScript(int maxActions)
    {
        var safeMaxActions = Math.Max(1, maxActions);

        return $@"
function Get-OptionalString($obj, [string]$propertyName) {{
    if ($null -eq $obj) {{ return $null }}
    $prop = $obj.PSObject.Properties[$propertyName]
    if ($null -eq $prop -or $null -eq $prop.Value) {{ return $null }}
    return $prop.Value.ToString()
}}

function Get-OptionalDate($obj, [string]$propertyName) {{
    $value = Get-OptionalString $obj $propertyName
    if ([string]::IsNullOrWhiteSpace($value)) {{ return $null }}
    try {{ return [datetime]$value }} catch {{ return $null }}
}}

function Get-StringList($value) {{
    $items = New-Object System.Collections.Generic.List[string]
    foreach ($item in @($value)) {{
        if ($null -eq $item) {{ continue }}
        $text = $item.ToString()
        if (-not [string]::IsNullOrWhiteSpace($text)) {{
            $items.Add($text)
        }}
    }}

    return @($items)
}}

function Write-WorkspaceWarning([string]$code, [string]$message, [string[]]$sampleItems, [bool]$isPartialData) {{
    $warningPayload = @{{
        Code = $code
        Scope = 'ComplianceWorkspace'
        Message = $message
        AffectedItemCount = if ($sampleItems.Count -gt 0) {{ $sampleItems.Count }} else {{ $null }}
        SampleItems = @($sampleItems | Select-Object -First 5)
        IsPartialData = $isPartialData
    }} | ConvertTo-Json -Compress -Depth 4
    Write-Warning ""__EA_WARN__$warningPayload""
}}

function Get-FallbackHoldEntries($searchItems) {{
    $holdIndex = @{{}}
    $candidateLocations = New-Object System.Collections.Generic.List[string]
    $skippedLocations = New-Object System.Collections.Generic.List[string]

    foreach ($searchItem in @($searchItems)) {{
        foreach ($location in @(Get-StringList $searchItem.ExchangeLocations)) {{
            if ([string]::IsNullOrWhiteSpace($location)) {{ continue }}
            if (-not $candidateLocations.Contains($location)) {{
                $candidateLocations.Add($location)
            }}
        }}
    }}

    foreach ($location in @($candidateLocations)) {{
        try {{
            $mailbox = Get-Mailbox -Identity $location -ErrorAction Stop | Select-Object -First 1
            foreach ($holdId in @(Get-StringList $mailbox.InPlaceHolds)) {{
                if (-not $holdId.StartsWith('UniH', [System.StringComparison]::OrdinalIgnoreCase)) {{
                    continue
                }}

                if (-not $holdIndex.ContainsKey($holdId)) {{
                    $holdIndex[$holdId] = [ordered]@{{
                        Identity = $holdId
                        Name = $holdId
                        ActionType = 'Hold'
                        SearchName = $null
                        CaseName = $null
                        Status = 'FallbackDetected'
                        CreatedBy = $null
                        CreatedTime = $null
                        CompletedTime = $null
                        ExchangeLocations = New-Object System.Collections.Generic.List[string]
                        Details = 'Fallback from mailbox InPlaceHolds: case/search metadata is not available without Get-CaseHoldPolicy.'
                    }}
                }}

                $entry = $holdIndex[$holdId]
                if (-not $entry.ExchangeLocations.Contains($location)) {{
                    $entry.ExchangeLocations.Add($location)
                }}
            }}
        }}
        catch {{
            if (-not $skippedLocations.Contains($location)) {{
                $skippedLocations.Add($location)
            }}
        }}
    }}

    $entries = @(
        $holdIndex.Values |
            ForEach-Object {{
                [PSCustomObject]@{{
                    Identity = $_.Identity
                    Name = $_.Name
                    ActionType = $_.ActionType
                    SearchName = $_.SearchName
                    CaseName = $_.CaseName
                    Status = $_.Status
                    CreatedBy = $_.CreatedBy
                    CreatedTime = $_.CreatedTime
                    CompletedTime = $_.CompletedTime
                    ExchangeLocations = @($_.ExchangeLocations)
                    Details = $_.Details
                }}
            }} |
            Sort-Object Name
    )

    return [PSCustomObject]@{{
        Entries = @($entries)
        CandidateLocations = @($candidateLocations)
        SkippedLocations = @($skippedLocations)
    }}
}}

$searches = @()
try {{
    $searches = @(
        Get-ComplianceSearch -ErrorAction Stop |
            Sort-Object Name |
            ForEach-Object {{
                [PSCustomObject]@{{
                    Name = Get-OptionalString $_ 'Name'
                    CaseName = Get-OptionalString $_ 'Case'
                    Status = Get-OptionalString $_ 'Status'
                    CreatedBy = Get-OptionalString $_ 'CreatedBy'
                    CreatedTime = Get-OptionalDate $_ 'CreatedTime'
                    LastModifiedTime = Get-OptionalDate $_ 'LastModifiedTime'
                    Items = Get-OptionalString $_ 'Items'
                    Size = Get-OptionalString $_ 'Size'
                    ExchangeLocations = @(Get-StringList $_.ExchangeLocation)
                    ContentMatchQuery = Get-OptionalString $_ 'ContentMatchQuery'
                }}
            }}
    )
}}
catch {{
    Write-WorkspaceWarning 'ComplianceSearchLoadFailed' ""Get-ComplianceSearch failed: $($_.Exception.Message)"" @() $true
}}

$cases = @()
try {{
    $cases = @(
        Get-ComplianceCase -ErrorAction Stop |
            Sort-Object Name |
            ForEach-Object {{
                [PSCustomObject]@{{
                    Name = Get-OptionalString $_ 'Name'
                    Status = Get-OptionalString $_ 'Status'
                    CaseType = Get-OptionalString $_ 'CaseType'
                    CreatedTime = Get-OptionalDate $_ 'CreatedTime'
                    Description = Get-OptionalString $_ 'Description'
                }}
            }}
    )
}}
catch {{
    Write-WorkspaceWarning 'ComplianceCaseLoadFailed' ""Get-ComplianceCase failed: $($_.Exception.Message)"" @() $true
}}

$purgeActions = @()
try {{
    $purgeActions = @(
        Get-ComplianceSearchAction -ErrorAction Stop |
            Sort-Object {{ Get-OptionalDate $_ 'CreatedTime' }} -Descending |
            Select-Object -First {safeMaxActions} |
            ForEach-Object {{
                [PSCustomObject]@{{
                    Identity = Get-OptionalString $_ 'Identity'
                    Name = Get-OptionalString $_ 'Name'
                    ActionType = if (Get-OptionalString $_ 'Action') {{ Get-OptionalString $_ 'Action' }} else {{ 'Action' }}
                    SearchName = Get-OptionalString $_ 'SearchName'
                    CaseName = $null
                    Status = Get-OptionalString $_ 'Status'
                    CreatedBy = Get-OptionalString $_ 'CreatedBy'
                    CreatedTime = Get-OptionalDate $_ 'CreatedTime'
                    CompletedTime = Get-OptionalDate $_ 'CompletedTime'
                    ExchangeLocations = @()
                    Details = if (Get-OptionalString $_ 'Results') {{ Get-OptionalString $_ 'Results' }} else {{ Get-OptionalString $_ 'Scenario' }}
                }}
            }}
    )
}}
catch {{
    Write-WorkspaceWarning 'ComplianceActionLoadFailed' ""Get-ComplianceSearchAction failed: $($_.Exception.Message)"" @() $true
}}

$holdActions = @()
if (-not (Get-Command -Name 'Get-CaseHoldPolicy' -ErrorAction SilentlyContinue)) {{
    $fallbackHoldReport = Get-FallbackHoldEntries -searchItems $searches
    $holdActions = @($fallbackHoldReport.Entries)

    if ($holdActions.Count -gt 0) {{
        Write-WorkspaceWarning 'CaseHoldFallbackUsed' 'Get-CaseHoldPolicy is not available. Visible holds were reconstructed from mailbox InPlaceHolds values; case/search metadata may be missing.' @($fallbackHoldReport.CandidateLocations) $true
    }}
    else {{
        Write-WorkspaceWarning 'CaseHoldPolicyUnavailable' 'Existing holds are not visible in this Purview session: Get-CaseHoldPolicy is not available and they could not be reconstructed from observable mailbox InPlaceHolds values.' @($fallbackHoldReport.CandidateLocations) $true
    }}

    if ($fallbackHoldReport.SkippedLocations.Count -gt 0) {{
        Write-WorkspaceWarning 'CaseHoldFallbackMailboxProbeFailed' ""Mailbox hold fallback failed for $($fallbackHoldReport.SkippedLocations.Count) location(s). The Compliance workspace remains partial."" @($fallbackHoldReport.SkippedLocations) $true
    }}
}}
else {{
    $holdActionFailures = New-Object System.Collections.Generic.List[string]
foreach ($caseItem in $cases) {{
    try {{
        $casePolicies = Get-CaseHoldPolicy -Case $caseItem.Name -DistributionDetail -ErrorAction Stop
        foreach ($policy in @($casePolicies)) {{
            $rule = $null
            try {{
                $rule = Get-CaseHoldRule -Policy $policy.Name -ErrorAction Stop | Select-Object -First 1
            }}
            catch {{
            }}

            $holdActions += [PSCustomObject]@{{
                Identity = Get-OptionalString $policy 'Identity'
                Name = $policy.Name
                ActionType = 'Hold'
                SearchName = $null
                CaseName = $caseItem.Name
                Status = if (Get-OptionalString $policy 'DistributionStatus') {{ Get-OptionalString $policy 'DistributionStatus' }} else {{ Get-OptionalString $policy 'Status' }}
                CreatedBy = Get-OptionalString $policy 'CreatedBy'
                CreatedTime = Get-OptionalDate $policy 'WhenCreatedUTC'
                CompletedTime = Get-OptionalDate $policy 'WhenChangedUTC'
                ExchangeLocations = @(Get-StringList $policy.ExchangeLocation)
                Details = if ($null -ne $rule) {{ Get-OptionalString $rule 'ContentMatchQuery' }} else {{ Get-OptionalString $policy 'Comment' }}
            }}
        }}
    }}
    catch {{
        $caseName = if ($caseItem.Name) {{ $caseItem.Name.ToString() }} else {{ '(unknown case)' }}
        $holdActionFailures.Add($caseName)
    }}
}}
    if ($holdActionFailures.Count -gt 0) {{
        Write-WorkspaceWarning 'CaseHoldLoadFailed' ""Get-CaseHoldPolicy failed for $($holdActionFailures.Count) case(s). The Compliance workspace remains partial."" @($holdActionFailures) $true
    }}
}}

$actions = @($purgeActions + $holdActions | Sort-Object CreatedTime -Descending | Select-Object -First {safeMaxActions})

[PSCustomObject]@{{
    Searches = @($searches)
    Cases = @($cases)
    Actions = @($actions)
}}";
    }

    internal static string BuildSearchUnifiedAuditLogScript(SearchUnifiedAuditLogRequest request)
    {
        var startDate = request.StartDate.ToString("O");
        var endDate = request.EndDate.ToString("O");
        var maxResults = Math.Max(1, request.MaxResults);
        var builder = new StringBuilder();
        builder.AppendLine("$params = @{");
        builder.AppendLine($"    StartDate = [datetime]'{EscapePs(startDate)}'");
        builder.AppendLine($"    EndDate = [datetime]'{EscapePs(endDate)}'");
        builder.AppendLine($"    ResultSize = {maxResults}");
        builder.AppendLine("}");

        if (request.UserIds.Count > 0)
        {
            builder.AppendLine($"$params['UserIds'] = {ToPsArrayLiteral(request.UserIds)}");
        }

        if (request.Operations.Count > 0)
        {
            builder.AppendLine($"$params['Operations'] = {ToPsArrayLiteral(request.Operations)}");
        }

        if (request.ObjectIds.Count > 0)
        {
            builder.AppendLine($"$params['ObjectIds'] = {ToPsArrayLiteral(request.ObjectIds)}");
        }

        if (!string.IsNullOrWhiteSpace(request.FreeText))
        {
            builder.AppendLine($"$params['FreeText'] = '{EscapePs(request.FreeText)}'");
        }

        builder.AppendLine(@"
Search-UnifiedAuditLog @params -ErrorAction Stop |
    Sort-Object CreationDate -Descending |
    ForEach-Object {
        [PSCustomObject]@{
            Identity = if ($_.Identity) { $_.Identity.ToString() } else { [guid]::NewGuid().ToString() }
            CreationDate = if ($_.CreationDate) { $_.CreationDate } else { $null }
            UserIds = if ($_.UserIds) { $_.UserIds.ToString() } else { $null }
            Operations = if ($_.Operations) { $_.Operations.ToString() } else { $null }
            RecordType = if ($_.RecordType) { $_.RecordType.ToString() } else { $null }
            ResultStatus = if ($_.ResultStatus) { $_.ResultStatus.ToString() } else { $null }
            ObjectId = if ($_.ObjectId) { $_.ObjectId.ToString() } else { $null }
            AuditData = if ($_.AuditData) { $_.AuditData.ToString() } else { $null }
        }
    }");

        return builder.ToString();
    }

    internal static SearchUnifiedAuditLogResponse MapSearchUnifiedAuditLogResponse(IEnumerable<PSObject> results, int maxResults)
    {
        var response = new SearchUnifiedAuditLogResponse();

        foreach (var obj in results)
        {
            response.Results.Add(new UnifiedAuditLogRecordDto
            {
                Identity = GetString(obj, "Identity"),
                CreationDate = GetNullableDateTime(obj, "CreationDate"),
                UserIds = GetNullableString(obj, "UserIds"),
                Operations = GetNullableString(obj, "Operations"),
                RecordType = GetNullableString(obj, "RecordType"),
                ResultStatus = GetNullableString(obj, "ResultStatus"),
                ObjectId = GetNullableString(obj, "ObjectId"),
                AuditData = GetNullableString(obj, "AuditData")
            });
        }

        response.TotalCount = response.Results.Count;
        if (maxResults > 0 && response.TotalCount >= maxResults)
        {
            response.Warning = $"Results limited to the first {maxResults} record(s) returned by Search-UnifiedAuditLog.";
        }

        return response;
    }

    private static ComplianceActionSummaryDto MapComplianceAction(PSObject obj)
    {
        return new ComplianceActionSummaryDto
        {
            Identity = GetString(obj, "Identity"),
            Name = GetString(obj, "Name"),
            ActionType = NormalizeActionType(GetNullableString(obj, "ActionType") ?? GetNullableString(obj, "Action") ?? "Action"),
            SearchName = GetNullableString(obj, "SearchName"),
            CaseName = GetNullableString(obj, "CaseName"),
            Status = GetNullableString(obj, "Status"),
            CreatedBy = GetNullableString(obj, "CreatedBy"),
            CreatedTime = GetNullableDateTime(obj, "CreatedTime"),
            CompletedTime = GetNullableDateTime(obj, "CompletedTime"),
            ExchangeLocations = ConvertToStringList(GetPropertyValue(obj, "ExchangeLocations")),
            Details = GetNullableString(obj, "Details") ?? GetNullableString(obj, "Results")
        };
    }

    private static ComplianceActionSummaryDto MapHoldAction(PSObject obj, string searchName, string? caseName)
    {
        return new ComplianceActionSummaryDto
        {
            Identity = GetString(obj, "Identity"),
            Name = GetString(obj, "Name"),
            ActionType = "Hold",
            SearchName = searchName,
            CaseName = caseName,
            Status = GetNullableString(obj, "DistributionStatus") ?? GetNullableString(obj, "Status"),
            CreatedBy = GetNullableString(obj, "CreatedBy"),
            CreatedTime = GetNullableDateTime(obj, "WhenCreatedUTC") ?? GetNullableDateTime(obj, "CreatedTime"),
            CompletedTime = GetNullableDateTime(obj, "WhenChangedUTC") ?? GetNullableDateTime(obj, "CompletedTime"),
            ExchangeLocations = ConvertToStringList(GetPropertyValue(obj, "ExchangeLocation")),
            Details = GetNullableString(obj, "Comment")
        };
    }

    private static ComplianceActionSummaryDto CreateFallbackAction(InvokeComplianceActionRequest request, ComplianceSearchDto search)
    {
        return new ComplianceActionSummaryDto
        {
            Identity = request.SearchName,
            Name = string.Equals(request.ActionType, "Hold", StringComparison.OrdinalIgnoreCase)
                ? request.HoldName ?? request.SearchName
                : $"{request.SearchName}_Purge",
            ActionType = NormalizeActionType(request.ActionType),
            SearchName = request.SearchName,
            CaseName = request.CaseName,
            Status = "Submitted",
            CreatedTime = DateTime.UtcNow,
            ExchangeLocations = search.ExchangeLocations,
            Details = search.ContentMatchQuery
        };
    }

    private static IEnumerable<PSObject> EnumeratePsObjects(object? value)
    {
        if (value is PSObject single)
        {
            yield return single;
            yield break;
        }

        if (value is IEnumerable enumerable)
        {
            foreach (var item in enumerable)
            {
                if (item is PSObject obj)
                {
                    yield return obj;
                }
                else if (item != null)
                {
                    yield return PSObject.AsPSObject(item);
                }
            }
        }
    }

    private static string NormalizeActionType(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "Action";
        }

        return value.Trim() switch
        {
            "Purge" => "Purge",
            "Hold" => "Hold",
            "Preview" => "Preview",
            _ => value.Trim()
        };
    }

    private static string NormalizePurgeType(string? value)
    {
        return value?.Trim() switch
        {
            "HardDelete" => "HardDelete",
            _ => "SoftDelete"
        };
    }
}
