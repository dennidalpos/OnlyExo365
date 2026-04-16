using System.Collections;
using System.Management.Automation;
using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

internal sealed class ExoResourceCommands : ExoCommandModuleBase
{
    private readonly ExoMailboxReportingCommands _mailboxReportingCommands;

    public ExoResourceCommands(
        PowerShellEngine engine,
        ExoMailboxReportingCommands mailboxReportingCommands)
        : base(engine)
    {
        _mailboxReportingCommands = mailboxReportingCommands;
    }

    public async Task<GetResourceMailboxesResponse> GetResourceMailboxesAsync(
        GetResourceMailboxesRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var normalizedType = NormalizeResourceType(request.ResourceType);
        var sortProperty = NormalizeResourceSortProperty(request.SortBy);
        var recipientTypeFilter = normalizedType switch
        {
            "Room" => "RoomMailbox",
            "Equipment" => "EquipmentMailbox",
            _ => "RoomMailbox,EquipmentMailbox"
        };

        var response = new GetResourceMailboxesResponse
        {
            Skip = request.Skip,
            PageSize = request.PageSize,
            SearchQuery = request.SearchQuery
        };

        if (!string.IsNullOrWhiteSpace(request.ResourceType) && normalizedType == null)
        {
            onLog?.Invoke("Warning", $"Unsupported ResourceType ignored: {request.ResourceType}");
        }

        if (!string.IsNullOrWhiteSpace(request.SortBy) &&
            !string.Equals(sortProperty, request.SortBy.Trim(), StringComparison.OrdinalIgnoreCase))
        {
            onLog?.Invoke("Warning", $"Unsupported SortBy ignored: {request.SortBy}");
        }

        var escapedSearch = EscapePs(request.SearchQuery);
        var sortDirection = request.SortDescending ? "-Descending" : string.Empty;

        var script = $@"
$items = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails {recipientTypeFilter} | ForEach-Object {{
    [PSCustomObject]@{{
        Identity = $_.Identity.ToString()
        DisplayName = $_.DisplayName
        Alias = $_.Alias
        PrimarySmtpAddress = if ($_.PrimarySmtpAddress) {{ $_.PrimarySmtpAddress.ToString() }} else {{ '' }}
        ResourceType = if ($_.RecipientTypeDetails.ToString() -eq 'EquipmentMailbox') {{ 'Equipment' }} else {{ 'Room' }}
        RecipientTypeDetails = $_.RecipientTypeDetails.ToString()
        HiddenFromAddressListsEnabled = [bool]$_.HiddenFromAddressListsEnabled
    }}
}}

$searchQuery = '{escapedSearch}'
if (-not [string]::IsNullOrWhiteSpace($searchQuery)) {{
    $items = $items | Where-Object {{
        $_.DisplayName -like ""*$searchQuery*"" -or
        $_.PrimarySmtpAddress -like ""*$searchQuery*"" -or
        $_.Alias -like ""*$searchQuery*""
    }}
}}

$items = $items | Sort-Object {sortProperty} {sortDirection}
$totalCount = @($items).Count
$pagedItems = $items | Select-Object -Skip {request.Skip} -First {request.PageSize}

@{{
    TotalCount = $totalCount
    Resources = @($pagedItems)
}}";

        onLog?.Invoke("Verbose", $"Fetching resource mailboxes (skip={request.Skip}, pageSize={request.PageSize}, type={normalizedType ?? "All"})...");

        var result = await Engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (result.Success && result.Output.Any() && result.Output.First().BaseObject is Hashtable hash)
        {
            response.TotalCount = Convert.ToInt32(hash["TotalCount"] ?? 0);

            if (hash["Resources"] is object[] resources)
            {
                foreach (var resourceObject in resources)
                {
                    if (resourceObject is not PSObject resourcePs)
                    {
                        continue;
                    }

                    response.Resources.Add(new ResourceMailboxListItemDto
                    {
                        Identity = GetString(resourcePs, "Identity"),
                        DisplayName = GetString(resourcePs, "DisplayName"),
                        Alias = GetString(resourcePs, "Alias"),
                        PrimarySmtpAddress = GetString(resourcePs, "PrimarySmtpAddress"),
                        ResourceType = GetString(resourcePs, "ResourceType"),
                        RecipientTypeDetails = GetString(resourcePs, "RecipientTypeDetails"),
                        HiddenFromAddressListsEnabled = GetBool(resourcePs, "HiddenFromAddressListsEnabled")
                    });
                }
            }

            response.HasMore = (request.Skip + response.Resources.Count) < response.TotalCount;
        }

        onLog?.Invoke("Information", $"Retrieved {response.Resources.Count} resource mailboxes (total: {response.TotalCount})");

        return response;
    }

    public async Task<ResourceMailboxDetailsDto> GetResourceMailboxDetailsAsync(
        GetResourceMailboxDetailsRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var escapedIdentity = EscapePs(request.Identity);

        var script = $@"
$mailbox = Get-Mailbox -Identity '{escapedIdentity}' -ErrorAction Stop
$calendar = Get-CalendarProcessing -Identity '{escapedIdentity}' -ErrorAction Stop

[PSCustomObject]@{{
    Identity = $mailbox.Identity.ToString()
    DisplayName = $mailbox.DisplayName
    Name = $mailbox.Name
    Alias = $mailbox.Alias
    PrimarySmtpAddress = if ($mailbox.PrimarySmtpAddress) {{ $mailbox.PrimarySmtpAddress.ToString() }} else {{ '' }}
    ResourceType = if ($mailbox.RecipientTypeDetails.ToString() -eq 'EquipmentMailbox') {{ 'Equipment' }} else {{ 'Room' }}
    RecipientTypeDetails = $mailbox.RecipientTypeDetails.ToString()
    HiddenFromAddressListsEnabled = [bool]$mailbox.HiddenFromAddressListsEnabled
    WhenCreated = $mailbox.WhenCreated
    AutomateProcessing = if ($calendar.AutomateProcessing) {{ $calendar.AutomateProcessing.ToString() }} else {{ 'AutoAccept' }}
    AllowConflicts = [bool]$calendar.AllowConflicts
    AllBookInPolicy = [bool]$calendar.AllBookInPolicy
    AllRequestInPolicy = [bool]$calendar.AllRequestInPolicy
    AllRequestOutOfPolicy = [bool]$calendar.AllRequestOutOfPolicy
    BookingWindowInDays = if ($null -ne $calendar.BookingWindowInDays) {{ [int]$calendar.BookingWindowInDays }} else {{ $null }}
    MaximumDurationInMinutes = if ($null -ne $calendar.MaximumDurationInMinutes) {{ [int]$calendar.MaximumDurationInMinutes }} else {{ $null }}
    DeleteSubject = [bool]$calendar.DeleteSubject
    AddOrganizerToSubject = [bool]$calendar.AddOrganizerToSubject
    RemovePrivateProperty = [bool]$calendar.RemovePrivateProperty
    EnforceSchedulingHorizon = [bool]$calendar.EnforceSchedulingHorizon
    BookInPolicy = @($calendar.BookInPolicy | ForEach-Object {{ $_.ToString() }})
    RequestInPolicy = @($calendar.RequestInPolicy | ForEach-Object {{ $_.ToString() }})
    RequestOutOfPolicy = @($calendar.RequestOutOfPolicy | ForEach-Object {{ $_.ToString() }})
    ResourceDelegates = @($calendar.ResourceDelegates | ForEach-Object {{ $_.ToString() }})
}}";

        onLog?.Invoke("Verbose", $"Fetching resource mailbox details for {request.Identity}...");

        var results = await RunScriptAsync(script, cancellationToken);
        if (results.Count == 0)
        {
            throw new InvalidOperationException($"Resource mailbox not found: {request.Identity}");
        }

        var obj = results[0];
        var details = new ResourceMailboxDetailsDto
        {
            Identity = GetString(obj, "Identity"),
            DisplayName = GetString(obj, "DisplayName"),
            Name = GetNullableString(obj, "Name"),
            Alias = GetString(obj, "Alias"),
            PrimarySmtpAddress = GetString(obj, "PrimarySmtpAddress"),
            ResourceType = GetString(obj, "ResourceType"),
            RecipientTypeDetails = GetString(obj, "RecipientTypeDetails"),
            HiddenFromAddressListsEnabled = GetBool(obj, "HiddenFromAddressListsEnabled"),
            WhenCreated = GetNullableDateTime(obj, "WhenCreated"),
            BookingSettings = new ResourceBookingSettingsDto
            {
                AutomateProcessing = GetString(obj, "AutomateProcessing"),
                AllowConflicts = GetBool(obj, "AllowConflicts"),
                AllBookInPolicy = GetBool(obj, "AllBookInPolicy"),
                AllRequestInPolicy = GetBool(obj, "AllRequestInPolicy"),
                AllRequestOutOfPolicy = GetBool(obj, "AllRequestOutOfPolicy"),
                BookingWindowInDays = GetNullableInt(obj, "BookingWindowInDays"),
                MaximumDurationInMinutes = GetNullableInt(obj, "MaximumDurationInMinutes"),
                DeleteSubject = GetNullableBool(obj, "DeleteSubject"),
                AddOrganizerToSubject = GetNullableBool(obj, "AddOrganizerToSubject"),
                RemovePrivateProperty = GetNullableBool(obj, "RemovePrivateProperty"),
                EnforceSchedulingHorizon = GetNullableBool(obj, "EnforceSchedulingHorizon"),
                BookInPolicy = ConvertToStringList(obj.Properties["BookInPolicy"]?.Value),
                RequestInPolicy = ConvertToStringList(obj.Properties["RequestInPolicy"]?.Value),
                RequestOutOfPolicy = ConvertToStringList(obj.Properties["RequestOutOfPolicy"]?.Value),
                ResourceDelegates = ConvertToStringList(obj.Properties["ResourceDelegates"]?.Value)
            }
        };

        details.Permissions = await _mailboxReportingCommands.GetMailboxPermissionsAsync(details.Identity, onLog, cancellationToken);

        return details;
    }

    public async Task<UpsertResourceMailboxResponse> UpsertResourceMailboxAsync(
        UpsertResourceMailboxRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var resourceType = NormalizeResourceType(request.ResourceType) ?? "Room";
        var identity = EscapePs(request.Identity);
        var displayName = EscapePs(request.DisplayName);
        var name = EscapePs(string.IsNullOrWhiteSpace(request.Name) ? request.DisplayName : request.Name);
        var alias = EscapePs(request.Alias);
        var primarySmtpAddress = EscapePs(request.PrimarySmtpAddress);
        var hidden = request.HiddenFromAddressListsEnabled ? "$true" : "$false";
        var booking = request.BookingSettings ?? new ResourceBookingSettingsDto();
        var bookingWindow = booking.BookingWindowInDays ?? 180;
        var maximumDuration = booking.MaximumDurationInMinutes ?? 1440;
        var automateProcessing = EscapePs(NormalizeAutomateProcessing(booking.AutomateProcessing));
        var bookInPolicy = ToPsArrayLiteral(booking.BookInPolicy);
        var requestInPolicy = ToPsArrayLiteral(booking.RequestInPolicy);
        var requestOutOfPolicy = ToPsArrayLiteral(booking.RequestOutOfPolicy);
        var resourceDelegates = ToPsArrayLiteral(booking.ResourceDelegates);

        string script;

        if (string.IsNullOrWhiteSpace(request.Identity))
        {
            var creationSwitch = resourceType == "Equipment" ? "-Equipment" : "-Room";
            script = $@"
New-Mailbox {creationSwitch} -Name '{name}' -DisplayName '{displayName}' -Alias '{alias}' -PrimarySmtpAddress '{primarySmtpAddress}' -ErrorAction Stop | Out-Null
Set-Mailbox -Identity '{primarySmtpAddress}' -HiddenFromAddressListsEnabled {hidden} -ErrorAction Stop
";
        }
        else
        {
            script = $@"
Set-Mailbox -Identity '{identity}' `
    -DisplayName '{displayName}' `
    -Name '{name}' `
    -Alias '{alias}' `
    -PrimarySmtpAddress '{primarySmtpAddress}' `
    -HiddenFromAddressListsEnabled {hidden} `
    -ErrorAction Stop
";
        }

        script += $@"
$bookInPolicy = {bookInPolicy}
$requestInPolicy = {requestInPolicy}
$requestOutOfPolicy = {requestOutOfPolicy}
$resourceDelegates = {resourceDelegates}
Set-CalendarProcessing -Identity '{primarySmtpAddress}' `
    -AutomateProcessing '{automateProcessing}' `
    -AllowConflicts {(booking.AllowConflicts ? "$true" : "$false")} `
    -AllBookInPolicy {(booking.AllBookInPolicy ? "$true" : "$false")} `
    -AllRequestInPolicy {(booking.AllRequestInPolicy ? "$true" : "$false")} `
    -AllRequestOutOfPolicy {(booking.AllRequestOutOfPolicy ? "$true" : "$false")} `
    -BookingWindowInDays {bookingWindow} `
    -MaximumDurationInMinutes {maximumDuration} `
    -DeleteSubject {ToPsBoolLiteral(booking.DeleteSubject ?? true)} `
    -AddOrganizerToSubject {ToPsBoolLiteral(booking.AddOrganizerToSubject ?? false)} `
    -RemovePrivateProperty {ToPsBoolLiteral(booking.RemovePrivateProperty ?? true)} `
    -EnforceSchedulingHorizon {ToPsBoolLiteral(booking.EnforceSchedulingHorizon ?? true)} `
    -BookInPolicy $bookInPolicy `
    -RequestInPolicy $requestInPolicy `
    -RequestOutOfPolicy $requestOutOfPolicy `
    -ResourceDelegates $resourceDelegates `
    -ErrorAction Stop

[PSCustomObject]@{{
    Identity = '{primarySmtpAddress}'
    PrimarySmtpAddress = '{primarySmtpAddress}'
}}";

        onLog?.Invoke("Information", $"Saving {resourceType} resource mailbox {request.PrimarySmtpAddress}...");

        var results = await RunScriptAsync(script, cancellationToken);
        var obj = results.LastOrDefault();

        return new UpsertResourceMailboxResponse
        {
            Identity = obj == null ? request.PrimarySmtpAddress : GetString(obj, "Identity"),
            PrimarySmtpAddress = obj == null ? request.PrimarySmtpAddress : GetString(obj, "PrimarySmtpAddress")
        };
    }

    private static string? NormalizeResourceType(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        return value.Trim() switch
        {
            "Room" => "Room",
            "RoomMailbox" => "Room",
            "Equipment" => "Equipment",
            "EquipmentMailbox" => "Equipment",
            _ => null
        };
    }

    private static string NormalizeResourceSortProperty(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "DisplayName";
        }

        return value.Trim() switch
        {
            "Alias" => "Alias",
            "PrimarySmtpAddress" => "PrimarySmtpAddress",
            "ResourceType" => "ResourceType",
            _ => "DisplayName"
        };
    }

    private static string NormalizeAutomateProcessing(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "AutoAccept";
        }

        return value.Trim() switch
        {
            "None" => "None",
            "AutoUpdate" => "AutoUpdate",
            _ => "AutoAccept"
        };
    }
}
