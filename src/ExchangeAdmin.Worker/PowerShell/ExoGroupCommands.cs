using System.Linq;
using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

             
                                                    
              
public class ExoGroupCommands
{
    private readonly PowerShellEngine _engine;
    private readonly CapabilityDetector _capabilityDetector;

    public ExoGroupCommands(PowerShellEngine engine, CapabilityDetector capabilityDetector)
    {
        _engine = engine;
        _capabilityDetector = capabilityDetector;
    }

    #region Distribution Lists

                 
                                           
                  
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

        var includeDynamic = request.IncludeDynamic;
        if (includeDynamic)
        {
            var capabilities = await _capabilityDetector.DetectCapabilitiesAsync(cancellationToken: cancellationToken);
            if (!capabilities.Features.CanGetDynamicDistributionGroup)
            {
                includeDynamic = false;
                onLog?.Invoke("Warning", "Get-DynamicDistributionGroup non disponibile: filtro liste dinamiche ignorato.");
            }
        }

                       
        var script = @"
$allGroups = @()

# Get distribution groups
$dgs = Get-DistributionGroup -ResultSize Unlimited
foreach ($dg in $dgs) {
    $allGroups += @{
        Type = 'DistributionGroup'
        Item = $dg
    }
}
";

                                          
        if (includeDynamic)
        {
            script += @"
# Get dynamic distribution groups
$ddgs = Get-DynamicDistributionGroup -ResultSize Unlimited
foreach ($ddg in $ddgs) {
    $allGroups += @{
        Type = 'DynamicDistributionGroup'
        Item = $ddg
    }
}
";
        }

                            
        if (!string.IsNullOrWhiteSpace(request.SearchQuery))
        {
            var escapedSearch = request.SearchQuery.Replace("'", "''");
            script += $@"
$allGroups = $allGroups | Where-Object {{
    $_.Item.DisplayName -like '*{escapedSearch}*' -or
    $_.Item.PrimarySmtpAddress -like '*{escapedSearch}*' -or
    $_.Item.Alias -like '*{escapedSearch}*'
}}
";
        }

               
        var sortProperty = string.IsNullOrWhiteSpace(request.SortBy) ? "DisplayName" : request.SortBy;
        var sortDirection = request.SortDescending ? "-Descending" : "";

        script += $@"
$allGroups = $allGroups | Sort-Object {{ $_.Item.{sortProperty} }} {sortDirection}
$totalCount = @($allGroups).Count

# Apply paging
$pagedGroups = $allGroups | Select-Object -Skip {request.Skip} -First {request.PageSize}

@{{
    TotalCount = $totalCount
    Groups = @($pagedGroups | ForEach-Object {{
        $g = $_.Item
        $isDynamic = $_.Type -eq 'DynamicDistributionGroup'
        @{{
            Identity = $g.Identity.ToString()
            Guid = $g.Guid.ToString()
            DisplayName = $g.DisplayName
            PrimarySmtpAddress = $g.PrimarySmtpAddress.ToString()
            Alias = $g.Alias
            GroupType = if ($isDynamic) {{ 'Dynamic' }} else {{ $g.GroupType.ToString() }}
            RecipientType = $g.RecipientType.ToString()
            RecipientTypeDetails = $g.RecipientTypeDetails.ToString()
            IsDynamic = $isDynamic
            ManagedBy = @($g.ManagedBy | ForEach-Object {{ $_.ToString() }})
        }}
    }})
}}
";

        onLog?.Invoke("Verbose", $"Fetching distribution lists (skip={request.Skip}, pageSize={request.PageSize})...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (result.Success && result.Output.Any())
        {
            var hash = result.Output.First().BaseObject as System.Collections.Hashtable;
            if (hash != null)
            {
                response.TotalCount = Convert.ToInt32(hash["TotalCount"] ?? 0);

                var groups = hash["Groups"] as object[];
                if (groups != null)
                {
                    foreach (var grpObj in groups)
                    {
                        if (grpObj is System.Collections.Hashtable grpHash)
                        {
                            var item = new DistributionListItemDto
                            {
                                Identity = grpHash["Identity"]?.ToString() ?? "",
                                Guid = grpHash["Guid"]?.ToString(),
                                DisplayName = grpHash["DisplayName"]?.ToString() ?? "",
                                PrimarySmtpAddress = grpHash["PrimarySmtpAddress"]?.ToString() ?? "",
                                Alias = grpHash["Alias"]?.ToString(),
                                GroupType = grpHash["GroupType"]?.ToString() ?? "",
                                RecipientType = grpHash["RecipientType"]?.ToString() ?? "",
                                RecipientTypeDetails = grpHash["RecipientTypeDetails"]?.ToString() ?? "",
                                IsDynamic = grpHash["IsDynamic"] as bool? ?? false,
                                ManagedBy = ConvertToStringList(grpHash["ManagedBy"])
                            };

                            response.DistributionLists.Add(item);
                            onPartialOutput?.Invoke(item);
                        }
                    }
                }

                response.HasMore = (request.Skip + response.DistributionLists.Count) < response.TotalCount;
            }
        }

        onLog?.Invoke("Information", $"Retrieved {response.DistributionLists.Count} distribution lists (total: {response.TotalCount})");

        return response;
    }

                 
                                       
                  
    public async Task<DistributionListDetailsDto> GetDistributionListDetailsAsync(
        GetDistributionListDetailsRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var escapedIdentity = request.Identity.Replace("'", "''");

                                                      
        var typeCheckScript = $@"
$dg = $null
$isDynamic = $false

try {{
    $dg = Get-DistributionGroup -Identity '{escapedIdentity}' -ErrorAction Stop
}}
catch {{
    try {{
        $dg = Get-DynamicDistributionGroup -Identity '{escapedIdentity}' -ErrorAction Stop
        $isDynamic = $true
    }}
    catch {{
        throw ""Distribution group not found: {escapedIdentity}""
    }}
}}

@{{ IsDynamic = $isDynamic; Group = $dg }}
";

        onLog?.Invoke("Verbose", $"Checking distribution list type for {request.Identity}...");

        var typeResult = await _engine.ExecuteAsync(typeCheckScript, onVerbose: onLog, cancellationToken: cancellationToken);

        if (typeResult.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!typeResult.Output.Any())
        {
            throw new InvalidOperationException($"Failed to get distribution list: {typeResult.ErrorMessage}");
        }

        var typeHash = typeResult.Output.First().BaseObject as System.Collections.Hashtable;
        var isDynamic = typeHash?["IsDynamic"] as bool? ?? false;

        if (isDynamic)
        {
            return await GetDynamicDistributionListDetailsAsync(request, onLog, cancellationToken);
        }

                                                
        var script = $@"
$dg = Get-DistributionGroup -Identity '{escapedIdentity}'

@{{
    Identity = $dg.Identity.ToString()
    Guid = $dg.Guid.ToString()
    DisplayName = $dg.DisplayName
    PrimarySmtpAddress = $dg.PrimarySmtpAddress.ToString()
    Alias = $dg.Alias
    GroupType = $dg.GroupType.ToString()
    RecipientType = $dg.RecipientType.ToString()
    RecipientTypeDetails = $dg.RecipientTypeDetails.ToString()
    EmailAddresses = @($dg.EmailAddresses | ForEach-Object {{ $_.ToString() }})
    ManagedBy = @($dg.ManagedBy | ForEach-Object {{ $_.ToString() }})
    AcceptMessagesOnlyFrom = @($dg.AcceptMessagesOnlyFrom | ForEach-Object {{ $_.ToString() }})
    RejectMessagesFrom = @($dg.RejectMessagesFrom | ForEach-Object {{ $_.ToString() }})
    RequireSenderAuthenticationEnabled = $dg.RequireSenderAuthenticationEnabled
    HiddenFromAddressListsEnabled = $dg.HiddenFromAddressListsEnabled
    MemberJoinRestriction = $dg.MemberJoinRestriction.ToString()
    MemberDepartRestriction = $dg.MemberDepartRestriction.ToString()
    WhenCreated = $dg.WhenCreated
    WhenChanged = $dg.WhenChanged
}}
";

        onLog?.Invoke("Verbose", $"Fetching distribution list details for {request.Identity}...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);

        if (!result.Output.Any())
        {
            throw new InvalidOperationException($"Failed to get distribution list: {result.ErrorMessage}");
        }

        var hash = result.Output.First().BaseObject as System.Collections.Hashtable;
        if (hash == null)
        {
            throw new InvalidOperationException("Failed to parse distribution list data");
        }

        var details = new DistributionListDetailsDto
        {
            Identity = hash["Identity"]?.ToString() ?? "",
            Guid = hash["Guid"]?.ToString(),
            DisplayName = hash["DisplayName"]?.ToString() ?? "",
            PrimarySmtpAddress = hash["PrimarySmtpAddress"]?.ToString() ?? "",
            Alias = hash["Alias"]?.ToString(),
            GroupType = hash["GroupType"]?.ToString() ?? "",
            RecipientType = hash["RecipientType"]?.ToString() ?? "",
            RecipientTypeDetails = hash["RecipientTypeDetails"]?.ToString() ?? "",
            EmailAddresses = ConvertToStringList(hash["EmailAddresses"]),
            ManagedBy = ConvertToStringList(hash["ManagedBy"]),
            AcceptMessagesOnlyFrom = ConvertToStringList(hash["AcceptMessagesOnlyFrom"]),
            RejectMessagesFrom = ConvertToStringList(hash["RejectMessagesFrom"]),
            RequireSenderAuthenticationEnabled = hash["RequireSenderAuthenticationEnabled"] as bool? ?? false,
            HiddenFromAddressListsEnabled = hash["HiddenFromAddressListsEnabled"] as bool? ?? false,
            MemberJoinRestriction = hash["MemberJoinRestriction"]?.ToString(),
            MemberDepartRestriction = hash["MemberDepartRestriction"]?.ToString(),
            WhenCreated = hash["WhenCreated"] as DateTime?,
            WhenChanged = hash["WhenChanged"] as DateTime?
        };

        cancellationToken.ThrowIfCancellationRequested();

                                   
        if (request.IncludeMembers)
        {
            details.Members = await GetGroupMembersPageAsync(
                request.Identity,
                "DistributionGroup",
                0,
                request.MembersPageSize,
                onLog,
                cancellationToken);
        }

        onLog?.Invoke("Information", $"Retrieved details for distribution list {details.DisplayName}");

        return details;
    }

    private async Task<DistributionListDetailsDto> GetDynamicDistributionListDetailsAsync(
        GetDistributionListDetailsRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var escapedIdentity = request.Identity.Replace("'", "''");

        var script = $@"
$ddg = Get-DynamicDistributionGroup -Identity '{escapedIdentity}'

@{{
    Identity = $ddg.Identity.ToString()
    Guid = $ddg.Guid.ToString()
    DisplayName = $ddg.DisplayName
    PrimarySmtpAddress = $ddg.PrimarySmtpAddress.ToString()
    Alias = $ddg.Alias
    RecipientFilter = $ddg.RecipientFilter
    IncludedRecipients = if ($ddg.IncludedRecipients) {{ $ddg.IncludedRecipients.ToString() }} else {{ $null }}
    ConditionalDepartment = @($ddg.ConditionalDepartment)
    ConditionalCompany = @($ddg.ConditionalCompany)
    ConditionalStateOrProvince = @($ddg.ConditionalStateOrProvince)
    ConditionalCustomAttribute1 = @($ddg.ConditionalCustomAttribute1)
    ManagedBy = @($ddg.ManagedBy | ForEach-Object {{ $_.ToString() }})
    EmailAddresses = @($ddg.EmailAddresses | ForEach-Object {{ $_.ToString() }})
    WhenCreated = $ddg.WhenCreated
    WhenChanged = $ddg.WhenChanged
}}
";

        onLog?.Invoke("Verbose", $"Fetching dynamic distribution list details for {request.Identity}...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);

        if (!result.Output.Any())
        {
            throw new InvalidOperationException($"Failed to get dynamic distribution list: {result.ErrorMessage}");
        }

        var hash = result.Output.First().BaseObject as System.Collections.Hashtable;
        if (hash == null)
        {
            throw new InvalidOperationException("Failed to parse dynamic distribution list data");
        }

                                                         
        var details = new DistributionListDetailsDto
        {
            Identity = hash["Identity"]?.ToString() ?? "",
            Guid = hash["Guid"]?.ToString(),
            DisplayName = hash["DisplayName"]?.ToString() ?? "",
            PrimarySmtpAddress = hash["PrimarySmtpAddress"]?.ToString() ?? "",
            Alias = hash["Alias"]?.ToString(),
            GroupType = "Dynamic",
            RecipientType = "DynamicDistributionGroup",
            RecipientTypeDetails = "DynamicDistributionGroup",
            EmailAddresses = ConvertToStringList(hash["EmailAddresses"]),
            ManagedBy = ConvertToStringList(hash["ManagedBy"]),
            WhenCreated = hash["WhenCreated"] as DateTime?,
            WhenChanged = hash["WhenChanged"] as DateTime?
        };

                                                                                 
                                  

        return details;
    }

                 
                                       
                  
    public async Task<GroupMembersPageDto> GetGroupMembersPageAsync(
        string identity,
        string groupType,
        int skip,
        int pageSize,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var escapedIdentity = identity.Replace("'", "''");

        string script;

        if (groupType == "DynamicDistributionGroup")
        {
                                                                     
            script = $@"
$ddg = Get-DynamicDistributionGroup -Identity '{escapedIdentity}'
$allMembers = Get-Recipient -RecipientPreviewFilter $ddg.RecipientFilter -ResultSize Unlimited

$totalCount = @($allMembers).Count
$pagedMembers = $allMembers | Select-Object -Skip {skip} -First {pageSize}

@{{
    TotalCount = $totalCount
    Members = @($pagedMembers | ForEach-Object {{
        @{{
            Identity = $_.Identity.ToString()
            Name = $_.DisplayName
            PrimarySmtpAddress = $_.PrimarySmtpAddress.ToString()
            RecipientType = $_.RecipientType.ToString()
        }}
    }})
}}
";
        }
        else
        {
                                         
            script = $@"
$allMembers = Get-DistributionGroupMember -Identity '{escapedIdentity}' -ResultSize Unlimited

$totalCount = @($allMembers).Count
$pagedMembers = $allMembers | Select-Object -Skip {skip} -First {pageSize}

@{{
    TotalCount = $totalCount
    Members = @($pagedMembers | ForEach-Object {{
        @{{
            Identity = $_.Identity.ToString()
            Name = $_.DisplayName
            PrimarySmtpAddress = if ($_.PrimarySmtpAddress) {{ $_.PrimarySmtpAddress.ToString() }} else {{ $null }}
            RecipientType = $_.RecipientType.ToString()
        }}
    }})
}}
";
        }

        onLog?.Invoke("Verbose", $"Fetching group members (skip={skip}, pageSize={pageSize})...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        var page = new GroupMembersPageDto
        {
            Skip = skip,
            PageSize = pageSize
        };

        if (result.Success && result.Output.Any())
        {
            var hash = result.Output.First().BaseObject as System.Collections.Hashtable;
            if (hash != null)
            {
                page.TotalCount = Convert.ToInt32(hash["TotalCount"] ?? 0);

                var members = hash["Members"] as object[];
                if (members != null)
                {
                    foreach (var memObj in members)
                    {
                        if (memObj is System.Collections.Hashtable memHash)
                        {
                            page.Members.Add(new GroupMemberDto
                            {
                                Identity = memHash["Identity"]?.ToString() ?? "",
                                Name = memHash["Name"]?.ToString() ?? "",
                                PrimarySmtpAddress = memHash["PrimarySmtpAddress"]?.ToString(),
                                RecipientType = memHash["RecipientType"]?.ToString()
                            });
                        }
                    }
                }

                page.HasMore = (skip + page.Members.Count) < page.TotalCount;
            }
        }

        onLog?.Invoke("Verbose", $"Retrieved {page.Members.Count} members (total: {page.TotalCount})");

        return page;
    }

                 
                                       
                  
    public async Task ModifyGroupMemberAsync(
        ModifyGroupMemberRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var escapedIdentity = request.Identity.Replace("'", "''");
        var escapedMember = request.Member.Replace("'", "''");

        string script;
        string actionVerb;

        if (request.Action == GroupMemberAction.Add)
        {
            actionVerb = "Adding";
            script = $"Add-DistributionGroupMember -Identity '{escapedIdentity}' -Member '{escapedMember}' -Confirm:$false";
        }
        else
        {
            actionVerb = "Removing";
            script = $"Remove-DistributionGroupMember -Identity '{escapedIdentity}' -Member '{escapedMember}' -Confirm:$false";
        }

        onLog?.Invoke("Information", $"{actionVerb} {request.Member} {(request.Action == GroupMemberAction.Add ? "to" : "from")} {request.Identity}...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to {request.Action} member: {result.ErrorMessage}");
        }

        onLog?.Invoke("Information", $"Successfully {(request.Action == GroupMemberAction.Add ? "added" : "removed")} member");
    }

                 
                                                    
                  
    public async Task CreateDistributionListAsync(
        CreateDistributionListRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(request.DisplayName) ||
            string.IsNullOrWhiteSpace(request.Alias) ||
            string.IsNullOrWhiteSpace(request.PrimarySmtpAddress))
        {
            throw new InvalidOperationException("DisplayName, Alias and PrimarySmtpAddress are required to create a distribution list.");
        }

        var escapedDisplayName = request.DisplayName.Replace("'", "''");
        var escapedAlias = request.Alias.Replace("'", "''");
        var escapedPrimarySmtpAddress = request.PrimarySmtpAddress.Replace("'", "''");

        var script = $"New-DistributionGroup -Name '{escapedDisplayName}' -DisplayName '{escapedDisplayName}' -Alias '{escapedAlias}' -PrimarySmtpAddress '{escapedPrimarySmtpAddress}'";

        onLog?.Invoke("Information", $"Creating distribution list {request.PrimarySmtpAddress}...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create distribution list: {result.ErrorMessage}");
        }

        onLog?.Invoke("Information", "Distribution list created successfully");
    }

    public async Task<PreviewDynamicGroupMembersResponse> PreviewDynamicGroupMembersAsync(
        PreviewDynamicGroupMembersRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var escapedIdentity = request.Identity.Replace("'", "''");

        var script = $@"
$ddg = Get-DynamicDistributionGroup -Identity '{escapedIdentity}'
$allMembers = Get-Recipient -RecipientPreviewFilter $ddg.RecipientFilter -ResultSize Unlimited

$totalCount = @($allMembers).Count
$isLimited = $totalCount -gt {request.MaxResults}
$previewMembers = $allMembers | Select-Object -First {request.MaxResults}

@{{
    TotalCount = $totalCount
    IsLimited = $isLimited
    Members = @($previewMembers | ForEach-Object {{
        @{{
            Identity = $_.Identity.ToString()
            Name = $_.DisplayName
            PrimarySmtpAddress = $_.PrimarySmtpAddress.ToString()
            RecipientType = $_.RecipientType.ToString()
        }}
    }})
}}
";

        onLog?.Invoke("Warning", $"Previewing dynamic group members (may be slow for large groups)...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        var response = new PreviewDynamicGroupMembersResponse
        {
            Identity = request.Identity
        };

        if (result.Success && result.Output.Any())
        {
            var hash = result.Output.First().BaseObject as System.Collections.Hashtable;
            if (hash != null)
            {
                response.TotalCount = Convert.ToInt32(hash["TotalCount"] ?? 0);
                response.IsLimited = hash["IsLimited"] as bool? ?? false;

                if (response.IsLimited)
                {
                    response.Warning = $"Preview limited to {request.MaxResults} of {response.TotalCount} members";
                }

                var members = hash["Members"] as object[];
                if (members != null)
                {
                    foreach (var memObj in members)
                    {
                        if (memObj is System.Collections.Hashtable memHash)
                        {
                            response.Members.Add(new GroupMemberDto
                            {
                                Identity = memHash["Identity"]?.ToString() ?? "",
                                Name = memHash["Name"]?.ToString() ?? "",
                                PrimarySmtpAddress = memHash["PrimarySmtpAddress"]?.ToString(),
                                RecipientType = memHash["RecipientType"]?.ToString()
                            });
                        }
                    }
                }
            }
        }

        onLog?.Invoke("Information", $"Preview complete: {response.Members.Count} members shown (total: {response.TotalCount})");

        return response;
    }

    public async Task SetDistributionListSettingsAsync(
        SetDistributionListSettingsRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
    {
        var escapedIdentity = request.Identity.Replace("'", "''");
        var setParams = new List<string>();

        var capabilities = await _capabilityDetector.DetectCapabilitiesAsync(cancellationToken: cancellationToken);

        if (request.RequireSenderAuthenticationEnabled.HasValue)
        {
            if (capabilities.Features.CanSetDistributionGroupRequireSenderAuthentication)
            {
                setParams.Add($"-RequireSenderAuthenticationEnabled ${request.RequireSenderAuthenticationEnabled.Value.ToString().ToLowerInvariant()}");
            }
            else
            {
                onLog?.Invoke("Warning", "Parametro RequireSenderAuthenticationEnabled non supportato: modifica ignorata.");
            }
        }

        if (request.AcceptMessagesOnlyFrom != null)
        {
            if (capabilities.Features.CanSetDistributionGroupAcceptMessagesOnlyFrom)
            {
                setParams.Add($"-AcceptMessagesOnlyFrom {FormatStringArrayParameter(request.AcceptMessagesOnlyFrom)}");
            }
            else
            {
                onLog?.Invoke("Warning", "Parametro AcceptMessagesOnlyFrom non supportato: modifica ignorata.");
            }
        }

        if (request.RejectMessagesFrom != null)
        {
            if (capabilities.Features.CanSetDistributionGroupRejectMessagesFrom)
            {
                setParams.Add($"-RejectMessagesFrom {FormatStringArrayParameter(request.RejectMessagesFrom)}");
            }
            else
            {
                onLog?.Invoke("Warning", "Parametro RejectMessagesFrom non supportato: modifica ignorata.");
            }
        }

        if (setParams.Count == 0)
        {
            return;
        }

        var script = $"Set-DistributionGroup -Identity '{escapedIdentity}' {string.Join(" ", setParams)}";

        onLog?.Invoke("Information", $"Updating distribution list settings for {request.Identity}...");

        var result = await _engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to update distribution list settings: {result.ErrorMessage}");
        }

        onLog?.Invoke("Information", "Distribution list settings updated successfully");
    }

    #endregion

    #region Helpers

    private static List<string> ConvertToStringList(object? obj)
    {
        if (obj == null) return new List<string>();

        if (obj is object[] array)
        {
            return array.Select(x => x?.ToString() ?? "").Where(x => !string.IsNullOrEmpty(x)).ToList();
        }

        if (obj is System.Collections.IEnumerable enumerable)
        {
            var list = new List<string>();
            foreach (var item in enumerable)
            {
                var str = item?.ToString();
                if (!string.IsNullOrEmpty(str))
                {
                    list.Add(str);
                }
            }
            return list;
        }

        var single = obj.ToString();
        return string.IsNullOrEmpty(single) ? new List<string>() : new List<string> { single };
    }

    private static string FormatStringArrayParameter(IEnumerable<string> values)
    {
        var sanitized = values
            .Select(value => value?.Trim())
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => $"'{value!.Replace("'", "''")}'")
            .ToList();

        if (sanitized.Count == 0)
        {
            return "$null";
        }

        return $"@({string.Join(", ", sanitized)})";
    }

    #endregion
}
