using System.Collections;
using System.Linq;
using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

public partial class ExoGroupCommands
{
    public async Task<GetDistributionListsResponse> GetDistributionListsAsync(
        GetDistributionListsRequest request,
        Action<string, string>? onLog = null,
        Action<DistributionListItemDto>? onPartialOutput = null,
        CancellationToken cancellationToken = default)
    {
        var response = new GetDistributionListsResponse
        {
            Skip = request.Skip,
            PageSize = request.PageSize,
            SearchQuery = request.SearchQuery
        };

        var capabilities = await _capabilityDetector.DetectCapabilitiesAsync(cancellationToken: cancellationToken);
        var includeDynamic = request.IncludeDynamic && capabilities.Features.CanGetDynamicDistributionGroup;
        var includeUnified = capabilities.Features.CanGetUnifiedGroup;

        if (request.IncludeDynamic && !includeDynamic)
        {
            onLog?.Invoke("Warning", "Get-DynamicDistributionGroup is not available: dynamic groups are excluded from the listing.");
        }

        var filterType = NormalizeGroupTypeFilter(request.GroupTypeFilter);
        if (filterType == GroupTypeDynamic && !includeDynamic)
        {
            filterType = null;
        }

        if (!string.IsNullOrWhiteSpace(request.Filter))
        {
            onLog?.Invoke("Warning", "Free-text group filtering is not supported by the hardened worker and will be ignored.");
        }

        var sortProperty = ExoRequestSanitizer.NormalizeGroupSortProperty(request.SortBy);
        if (!string.IsNullOrWhiteSpace(request.SortBy) &&
            !string.Equals(sortProperty, request.SortBy.Trim(), StringComparison.OrdinalIgnoreCase))
        {
            onLog?.Invoke("Warning", $"Unsupported SortBy ignored: {request.SortBy}");
        }

        var escapedSearch = EscapePowerShellString(request.SearchQuery);
        var escapedFilterType = EscapePowerShellString(filterType);
        var sortDirection = request.SortDescending ? "-Descending" : string.Empty;
        var useWindowedLoad = false;
        var script = BuildGetDistributionListsScript(
            request.Skip,
            request.PageSize,
            escapedSearch,
            escapedFilterType,
            sortProperty,
            sortDirection,
            includeDynamic,
            includeUnified,
            useWindowedLoad);

        onLog?.Invoke("Verbose", $"Fetching groups (skip={request.Skip}, pageSize={request.PageSize}, filter={filterType ?? "All"}, mode=full)...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (result.Success && result.Output.Any() && result.Output.First().BaseObject is Hashtable hash)
        {
            response.TotalCount = Convert.ToInt32(hash["TotalCount"] ?? 0);
            response.IsTotalCountExact = hash["IsTotalCountExact"] as bool? ?? false;
            response.HasMore = hash["HasMore"] as bool? ?? false;

            if (hash["Groups"] is object[] groups)
            {
                foreach (var groupObject in groups.OfType<Hashtable>())
                {
                    var item = new DistributionListItemDto
                    {
                        Identity = groupObject["Identity"]?.ToString() ?? string.Empty,
                        Guid = groupObject["Guid"]?.ToString(),
                        DisplayName = groupObject["DisplayName"]?.ToString() ?? string.Empty,
                        PrimarySmtpAddress = groupObject["PrimarySmtpAddress"]?.ToString() ?? string.Empty,
                        Alias = groupObject["Alias"]?.ToString(),
                        GroupType = groupObject["GroupType"]?.ToString() ?? GroupTypeDistribution,
                        RecipientType = groupObject["RecipientType"]?.ToString() ?? string.Empty,
                        RecipientTypeDetails = groupObject["RecipientTypeDetails"]?.ToString() ?? string.Empty,
                        IsDynamic = groupObject["IsDynamic"] as bool? ?? false,
                        ManagedBy = ConvertToStringList(groupObject["ManagedBy"]),
                        MemberCount = groupObject["MemberCount"] as int?
                    };

                    response.DistributionLists.Add(item);
                    onPartialOutput?.Invoke(item);
                }
            }
        }

        onLog?.Invoke("Information", $"Retrieved {response.DistributionLists.Count} groups (total: {response.TotalCount})");
        return response;
    }

    internal static string BuildGetDistributionListsScript(
        int skip,
        int pageSize,
        string escapedSearch,
        string escapedFilterType,
        string sortProperty,
        string sortDirection,
        bool includeDynamic,
        bool includeUnified,
        bool useWindowedLoad)
    {
        var resultSizeExpression = useWindowedLoad
            ? "$pageWindowSize"
            : "Unlimited";
        var pageWindowSizeDeclaration = useWindowedLoad
            ? $"$pageWindowSize = {CalculatePageWindowSize(skip, pageSize)}{Environment.NewLine}"
            : string.Empty;
        var exactFlag = useWindowedLoad ? "$false" : "$true";
        var totalCountExpression = useWindowedLoad
            ? $"{skip} + @($pagedGroups).Count + $(if ($hasMore) {{ 1 }} else {{ 0 }})"
            : "@($allGroups).Count";
        var hasMoreExpression = useWindowedLoad
            ? "@($allGroups).Count -gt " + pageSize
            : $"({skip} + @($pagedGroups).Count) -lt $totalCount";
        var searchFilter = string.IsNullOrWhiteSpace(escapedSearch)
            ? string.Empty
            : $@"
if ('{escapedSearch}') {{
    $allGroups = $allGroups | Where-Object {{
        $_.Item.DisplayName -like '*{escapedSearch}*' -or
        $_.Item.PrimarySmtpAddress -like '*{escapedSearch}*' -or
        $_.Item.Alias -like '*{escapedSearch}*'
    }}
}}";
        var sortBlock = useWindowedLoad
            ? string.Empty
            : $@"
$allGroups = $allGroups | Sort-Object {{ $_.Item.{sortProperty} }} {sortDirection}
$totalCount = @($allGroups).Count";

        return $@"
{pageWindowSizeDeclaration}$allGroups = @()

$distributionGroups = Get-DistributionGroup -ResultSize {resultSizeExpression}
foreach ($group in $distributionGroups) {{
    $groupType = if ($group.RecipientTypeDetails -eq 'MailUniversalSecurityGroup') {{ '{GroupTypeMailSecurity}' }} else {{ '{GroupTypeDistribution}' }}
    $allGroups += [pscustomobject]@{{
        GroupType = $groupType
        Item = $group
    }}
}}

if ({(includeDynamic ? "$true" : "$false")}) {{
    $dynamicGroups = Get-DynamicDistributionGroup -ResultSize {resultSizeExpression}
    foreach ($group in $dynamicGroups) {{
        $allGroups += [pscustomobject]@{{
            GroupType = '{GroupTypeDynamic}'
            Item = $group
        }}
    }}
}}

if ({(includeUnified ? "$true" : "$false")}) {{
    $unifiedGroups = Get-UnifiedGroup -ResultSize {resultSizeExpression}
    foreach ($group in $unifiedGroups) {{
        $allGroups += [pscustomobject]@{{
            GroupType = '{GroupTypeMicrosoft365}'
            Item = $group
        }}
    }}
}}

if ('{escapedFilterType}') {{
    $allGroups = $allGroups | Where-Object {{ $_.GroupType -eq '{escapedFilterType}' }}
}}
{searchFilter}{sortBlock}
$pagedGroups = $allGroups | Select-Object -Skip {skip} -First {pageSize}
$hasMore = {hasMoreExpression}

@{{
    TotalCount = {totalCountExpression}
    HasMore = $hasMore
    IsTotalCountExact = {exactFlag}
    Groups = @($pagedGroups | ForEach-Object {{
        $group = $_.Item
        $groupType = $_.GroupType
        @{{
            Identity = $group.Identity.ToString()
            Guid = if ($group.Guid) {{ $group.Guid.ToString() }} else {{ $null }}
            DisplayName = $group.DisplayName
            PrimarySmtpAddress = if ($group.PrimarySmtpAddress) {{ $group.PrimarySmtpAddress.ToString() }} else {{ '' }}
            Alias = $group.Alias
            GroupType = $groupType
            RecipientType = if ($group.RecipientType) {{ $group.RecipientType.ToString() }} else {{ '' }}
            RecipientTypeDetails = if ($group.RecipientTypeDetails) {{ $group.RecipientTypeDetails.ToString() }} else {{ '' }}
            IsDynamic = $groupType -eq '{GroupTypeDynamic}'
            ManagedBy = @($group.ManagedBy | ForEach-Object {{ $_.ToString() }})
            MemberCount = $null
        }}
    }})
}}";
    }

    internal static int CalculatePageWindowSize(int skip, int pageSize)
        => Math.Max(0, skip) + Math.Max(1, pageSize) + 1;
}
