using System.Collections;
using OnlyExo365.Contracts.Dtos;

namespace OnlyExo365.Worker.PowerShell;

internal sealed class ExoMailboxListingCommands : ExoCommandModuleBase
{
    public ExoMailboxListingCommands(PowerShellEngine engine)
        : base(engine)
    {
    }

    public async Task<GetMailboxesResponse> GetMailboxesAsync(
        GetMailboxesRequest request,
        Action<string, string>? onLog = null,
        Action<MailboxListItemDto>? onPartialOutput = null,
        CancellationToken cancellationToken = default)
    {
        var response = new GetMailboxesResponse
        {
            Skip = request.Skip,
            PageSize = request.PageSize,
            SearchQuery = request.SearchQuery
        };

        var filterParts = new List<string>();
        var recipientTypeDetails = ExoRequestSanitizer.NormalizeMailboxRecipientTypeDetails(request.RecipientTypeDetails);
        var sortProperty = ExoRequestSanitizer.NormalizeMailboxSortProperty(request.SortBy);

        if (!string.IsNullOrWhiteSpace(request.RecipientTypeDetails) && recipientTypeDetails == null)
        {
            onLog?.Invoke("Warning", $"Unsupported RecipientTypeDetails ignored: {request.RecipientTypeDetails}");
        }

        if (recipientTypeDetails != null)
        {
            filterParts.Add($"RecipientTypeDetails -eq '{recipientTypeDetails}'");
        }

        if (!string.IsNullOrWhiteSpace(request.Filter))
        {
            onLog?.Invoke("Warning", "Free-text mailbox filtering is not supported by the hardened worker and will be ignored.");
        }

        var filterParam = filterParts.Count > 0
            ? $"-Filter \"{string.Join(" -and ", filterParts)}\""
            : string.Empty;
        var escapedSearch = request.SearchQuery?.Replace("'", "''") ?? string.Empty;
        var useWindowedLoad = false;

        if (!string.IsNullOrWhiteSpace(request.SortBy) &&
            !string.Equals(sortProperty, request.SortBy.Trim(), StringComparison.OrdinalIgnoreCase))
        {
            onLog?.Invoke("Warning", $"Unsupported SortBy ignored: {request.SortBy}");
        }

        var sortDirection = request.SortDescending ? "-Descending" : string.Empty;
        var script = ExoMailboxScriptFactory.BuildGetMailboxesScript(
            request.Skip,
            request.PageSize,
            filterParam,
            escapedSearch,
            sortProperty,
            sortDirection,
            useWindowedLoad);

        onLog?.Invoke("Verbose", $"Fetching mailboxes (skip={request.Skip}, pageSize={request.PageSize}, mode=full)...");

        var result = await Engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (result.Success && result.Output.Any() && result.Output.First().BaseObject is Hashtable hash)
        {
            response.TotalCount = Convert.ToInt32(hash["TotalCount"] ?? 0);
            response.IsTotalCountExact = hash["IsTotalCountExact"] as bool? ?? false;
            response.HasMore = hash["HasMore"] as bool? ?? false;

            if (hash["Mailboxes"] is object[] mailboxes)
            {
                foreach (var mailboxHash in mailboxes.OfType<Hashtable>())
                {
                    var item = ExoMailboxMapper.ToMailboxListItem(mailboxHash);
                    response.Mailboxes.Add(item);
                    onPartialOutput?.Invoke(item);
                }
            }
        }

        onLog?.Invoke("Information", $"Retrieved {response.Mailboxes.Count} mailboxes (total: {response.TotalCount})");
        return response;
    }

    public async Task<GetDeletedMailboxesResponse> GetDeletedMailboxesAsync(
        GetDeletedMailboxesRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var response = new GetDeletedMailboxesResponse
        {
            Skip = request.Skip,
            PageSize = request.PageSize,
            SearchQuery = request.SearchQuery
        };

        var script = ExoMailboxScriptFactory.BuildGetDeletedMailboxesScript(request);

        onLog?.Invoke("Verbose", $"Fetching deleted mailboxes (skip={request.Skip}, pageSize={request.PageSize})...");

        var result = await Engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (result.Success && result.Output.Any() && result.Output.First().BaseObject is Hashtable hash)
        {
            response.TotalCount = Convert.ToInt32(hash["TotalCount"] ?? 0);

            if (hash["Mailboxes"] is object[] mailboxes)
            {
                foreach (var mailboxHash in mailboxes.OfType<Hashtable>())
                {
                    response.Mailboxes.Add(ExoMailboxMapper.ToDeletedMailboxItem(mailboxHash));
                }
            }

            response.HasMore = (request.Skip + response.Mailboxes.Count) < response.TotalCount;
        }

        onLog?.Invoke("Information", $"Retrieved {response.Mailboxes.Count} deleted mailboxes (total: {response.TotalCount})");
        return response;
    }
}

