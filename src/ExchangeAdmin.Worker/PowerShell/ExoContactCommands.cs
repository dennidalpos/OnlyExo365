using System.Collections;
using System.Management.Automation;
using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

internal sealed class ExoContactCommands : ExoCommandModuleBase
{
    public ExoContactCommands(PowerShellEngine engine)
        : base(engine)
    {
    }

    public async Task<GetContactsResponse> GetContactsAsync(
        GetContactsRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var normalizedKind = NormalizeContactKind(request.ContactKind);
        var sortProperty = NormalizeContactSortProperty(request.SortBy);

        var response = new GetContactsResponse
        {
            Skip = request.Skip,
            PageSize = request.PageSize,
            SearchQuery = request.SearchQuery
        };

        if (!string.IsNullOrWhiteSpace(request.ContactKind) && normalizedKind == null)
        {
            onLog?.Invoke("Warning", $"Unsupported ContactKind ignored: {request.ContactKind}");
        }

        if (!string.IsNullOrWhiteSpace(request.SortBy) &&
            !string.Equals(sortProperty, request.SortBy.Trim(), StringComparison.OrdinalIgnoreCase))
        {
            onLog?.Invoke("Warning", $"Unsupported SortBy ignored: {request.SortBy}");
        }

        var escapedSearch = EscapePs(request.SearchQuery);
        var sortDirection = request.SortDescending ? "-Descending" : string.Empty;
        var includeContacts = normalizedKind == null || string.Equals(normalizedKind, "MailContact", StringComparison.OrdinalIgnoreCase);
        var includeUsers = normalizedKind == null || string.Equals(normalizedKind, "MailUser", StringComparison.OrdinalIgnoreCase);

        var script = $@"
$items = @()

if ({(includeContacts ? "$true" : "$false")}) {{
    try {{
        $items += Get-MailContact -ResultSize Unlimited | ForEach-Object {{
            [PSCustomObject]@{{
                Identity = $_.Identity.ToString()
                Guid = if ($_.Guid) {{ $_.Guid.ToString() }} else {{ $null }}
                DisplayName = $_.DisplayName
                Name = $_.Name
                Alias = $_.Alias
                PrimarySmtpAddress = if ($_.PrimarySmtpAddress) {{ $_.PrimarySmtpAddress.ToString() }} else {{ '' }}
                ExternalEmailAddress = if ($_.ExternalEmailAddress) {{ $_.ExternalEmailAddress.ToString() }} else {{ '' }}
                RecipientType = $_.RecipientType.ToString()
                RecipientTypeDetails = $_.RecipientTypeDetails.ToString()
                ContactKind = 'MailContact'
                UserPrincipalName = $null
                HiddenFromAddressListsEnabled = [bool]$_.HiddenFromAddressListsEnabled
            }}
        }}
    }}
    catch {{
        Write-Warning ""Get-MailContact failed: $($_.Exception.Message)""
    }}
}}

if ({(includeUsers ? "$true" : "$false")}) {{
    try {{
        $items += Get-MailUser -ResultSize Unlimited | ForEach-Object {{
            [PSCustomObject]@{{
                Identity = $_.Identity.ToString()
                Guid = if ($_.Guid) {{ $_.Guid.ToString() }} else {{ $null }}
                DisplayName = $_.DisplayName
                Name = $_.Name
                Alias = $_.Alias
                PrimarySmtpAddress = if ($_.PrimarySmtpAddress) {{ $_.PrimarySmtpAddress.ToString() }} else {{ '' }}
                ExternalEmailAddress = if ($_.ExternalEmailAddress) {{ $_.ExternalEmailAddress.ToString() }} else {{ '' }}
                RecipientType = $_.RecipientType.ToString()
                RecipientTypeDetails = $_.RecipientTypeDetails.ToString()
                ContactKind = 'MailUser'
                UserPrincipalName = $_.MicrosoftOnlineServicesID
                HiddenFromAddressListsEnabled = [bool]$_.HiddenFromAddressListsEnabled
            }}
        }}
    }}
    catch {{
        Write-Warning ""Get-MailUser failed: $($_.Exception.Message)""
    }}
}}

$searchQuery = '{escapedSearch}'
if (-not [string]::IsNullOrWhiteSpace($searchQuery)) {{
    $items = $items | Where-Object {{
        $_.DisplayName -like ""*$searchQuery*"" -or
        $_.PrimarySmtpAddress -like ""*$searchQuery*"" -or
        $_.ExternalEmailAddress -like ""*$searchQuery*"" -or
        $_.Alias -like ""*$searchQuery*"" -or
        $_.UserPrincipalName -like ""*$searchQuery*""
    }}
}}

$items = $items | Sort-Object {sortProperty} {sortDirection}
$totalCount = @($items).Count
$pagedItems = $items | Select-Object -Skip {request.Skip} -First {request.PageSize}

@{{
    TotalCount = $totalCount
    Contacts = @($pagedItems)
}}";

        onLog?.Invoke("Verbose", $"Fetching contacts (skip={request.Skip}, pageSize={request.PageSize}, kind={normalizedKind ?? "All"})...");

        var result = await Engine.ExecuteAsync(script, onVerbose: onLog, cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (result.Success && result.Output.Any() && result.Output.First().BaseObject is Hashtable hash)
        {
            response.TotalCount = Convert.ToInt32(hash["TotalCount"] ?? 0);

            if (hash["Contacts"] is object[] contacts)
            {
                foreach (var contactObject in contacts)
                {
                    if (contactObject is not PSObject contactPs)
                    {
                        continue;
                    }

                    response.Contacts.Add(new ContactListItemDto
                    {
                        Identity = GetString(contactPs, "Identity"),
                        Guid = GetNullableString(contactPs, "Guid"),
                        DisplayName = GetString(contactPs, "DisplayName"),
                        Name = GetNullableString(contactPs, "Name"),
                        Alias = GetNullableString(contactPs, "Alias"),
                        PrimarySmtpAddress = GetString(contactPs, "PrimarySmtpAddress"),
                        ExternalEmailAddress = GetNullableString(contactPs, "ExternalEmailAddress"),
                        RecipientType = GetString(contactPs, "RecipientType"),
                        RecipientTypeDetails = GetString(contactPs, "RecipientTypeDetails"),
                        ContactKind = GetString(contactPs, "ContactKind"),
                        UserPrincipalName = GetNullableString(contactPs, "UserPrincipalName"),
                        HiddenFromAddressListsEnabled = GetBool(contactPs, "HiddenFromAddressListsEnabled")
                    });
                }
            }

            response.HasMore = (request.Skip + response.Contacts.Count) < response.TotalCount;
        }

        onLog?.Invoke("Information", $"Retrieved {response.Contacts.Count} contacts (total: {response.TotalCount})");

        return response;
    }

    public async Task<ContactDetailsDto> GetContactDetailsAsync(
        GetContactDetailsRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        var normalizedKind = NormalizeContactKind(request.ContactKind);
        var escapedIdentity = EscapePs(request.Identity);

        var script = $@"
$identity = '{escapedIdentity}'
$kind = '{EscapePs(normalizedKind)}'

function Get-ContactRecord($recipient) {{
    [PSCustomObject]@{{
        Identity = $recipient.Identity.ToString()
        Guid = if ($recipient.Guid) {{ $recipient.Guid.ToString() }} else {{ $null }}
        DisplayName = $recipient.DisplayName
        Name = $recipient.Name
        Alias = $recipient.Alias
        PrimarySmtpAddress = if ($recipient.PrimarySmtpAddress) {{ $recipient.PrimarySmtpAddress.ToString() }} else {{ '' }}
        ExternalEmailAddress = if ($recipient.ExternalEmailAddress) {{ $recipient.ExternalEmailAddress.ToString() }} else {{ '' }}
        RecipientType = $recipient.RecipientType.ToString()
        RecipientTypeDetails = $recipient.RecipientTypeDetails.ToString()
        ContactKind = if ($recipient.RecipientTypeDetails.ToString() -eq 'MailUser') {{ 'MailUser' }} else {{ 'MailContact' }}
        UserPrincipalName = if ($recipient.MicrosoftOnlineServicesID) {{ $recipient.MicrosoftOnlineServicesID.ToString() }} else {{ $null }}
        HiddenFromAddressListsEnabled = [bool]$recipient.HiddenFromAddressListsEnabled
        WhenCreated = $recipient.WhenCreated
        EmailAddresses = @($recipient.EmailAddresses | ForEach-Object {{ $_.ToString() }})
    }}
}}

if ($kind -eq 'MailUser') {{
    $recipient = Get-MailUser -Identity $identity -ErrorAction Stop
    Get-ContactRecord $recipient
}} elseif ($kind -eq 'MailContact') {{
    $recipient = Get-MailContact -Identity $identity -ErrorAction Stop
    Get-ContactRecord $recipient
}} else {{
    try {{
        $recipient = Get-MailContact -Identity $identity -ErrorAction Stop
        Get-ContactRecord $recipient
    }}
    catch {{
        $recipient = Get-MailUser -Identity $identity -ErrorAction Stop
        Get-ContactRecord $recipient
    }}
}}";

        onLog?.Invoke("Verbose", $"Fetching contact details for {request.Identity}...");

        var results = await RunScriptAsync(script, cancellationToken);
        if (results.Count == 0)
        {
            throw new InvalidOperationException($"Contact not found: {request.Identity}");
        }

        var obj = results[0];
        return new ContactDetailsDto
        {
            Identity = GetString(obj, "Identity"),
            Guid = GetNullableString(obj, "Guid"),
            DisplayName = GetString(obj, "DisplayName"),
            Name = GetNullableString(obj, "Name"),
            Alias = GetNullableString(obj, "Alias"),
            PrimarySmtpAddress = GetString(obj, "PrimarySmtpAddress"),
            ExternalEmailAddress = GetNullableString(obj, "ExternalEmailAddress"),
            RecipientType = GetString(obj, "RecipientType"),
            RecipientTypeDetails = GetString(obj, "RecipientTypeDetails"),
            ContactKind = GetString(obj, "ContactKind"),
            UserPrincipalName = GetNullableString(obj, "UserPrincipalName"),
            HiddenFromAddressListsEnabled = GetBool(obj, "HiddenFromAddressListsEnabled"),
            WhenCreated = GetNullableDateTime(obj, "WhenCreated"),
            EmailAddresses = ConvertToStringList(obj.Properties["EmailAddresses"]?.Value)
        };
    }

    public async Task UpsertContactAsync(UpsertContactRequest request, CancellationToken cancellationToken = default)
    {
        var command = BuildUpsertContactCommand(request);
        await RunScriptAsync(command.Script, command.Parameters, cancellationToken);
    }

    internal static (string Script, Dictionary<string, object>? Parameters) BuildUpsertContactCommand(UpsertContactRequest request)
    {
        var contactKind = NormalizeContactKind(request.ContactKind) ?? "MailContact";
        var identity = EscapePs(request.Identity);
        var displayName = EscapePs(request.DisplayName);
        var name = EscapePs(string.IsNullOrWhiteSpace(request.Name) ? request.DisplayName : request.Name);
        var alias = EscapePs(request.Alias);
        var primarySmtpAddress = EscapePs(request.PrimarySmtpAddress);
        var externalEmailAddress = EscapePs(request.ExternalEmailAddress);
        var userPrincipalName = EscapePs(request.UserPrincipalName);
        var hidden = request.HiddenFromAddressListsEnabled ? "$true" : "$false";

        if (contactKind.Equals("MailUser", StringComparison.OrdinalIgnoreCase))
        {
            var mailUserScript = $@"
$hidden = {hidden}
if ('{identity}' -ne '') {{
    Set-MailUser -Identity '{identity}' `
        -DisplayName '{displayName}' `
        -Name '{name}' `
        -Alias '{alias}' `
        -PrimarySmtpAddress '{primarySmtpAddress}' `
        -ExternalEmailAddress '{externalEmailAddress}' `
        -HiddenFromAddressListsEnabled $hidden `
        -ErrorAction Stop
}}
else {{
    $securePassword = ConvertTo-SecureString $PlainTextPassword -AsPlainText -Force
    New-MailUser -Name '{name}' `
        -DisplayName '{displayName}' `
        -Alias '{alias}' `
        -PrimarySmtpAddress '{primarySmtpAddress}' `
        -ExternalEmailAddress '{externalEmailAddress}' `
        -MicrosoftOnlineServicesID '{userPrincipalName}' `
        -Password $securePassword `
        -ErrorAction Stop
    if ($hidden) {{
        Set-MailUser -Identity '{primarySmtpAddress}' -HiddenFromAddressListsEnabled $hidden -ErrorAction Stop
    }}
}}
Write-Output 'OK'";

            return (mailUserScript, new Dictionary<string, object>
            {
                ["PlainTextPassword"] = request.Password ?? string.Empty
            });
        }

        var contactScript = $@"
$hidden = {hidden}
if ('{identity}' -ne '') {{
    Set-MailContact -Identity '{identity}' `
        -DisplayName '{displayName}' `
        -Name '{name}' `
        -Alias '{alias}' `
        -PrimarySmtpAddress '{primarySmtpAddress}' `
        -ExternalEmailAddress '{externalEmailAddress}' `
        -HiddenFromAddressListsEnabled $hidden `
        -ErrorAction Stop
}}
else {{
    New-MailContact -Name '{name}' `
        -DisplayName '{displayName}' `
        -Alias '{alias}' `
        -PrimarySmtpAddress '{primarySmtpAddress}' `
        -ExternalEmailAddress '{externalEmailAddress}' `
        -ErrorAction Stop
    if ($hidden) {{
        Set-MailContact -Identity '{primarySmtpAddress}' -HiddenFromAddressListsEnabled $hidden -ErrorAction Stop
    }}
}}
Write-Output 'OK'";

        return (contactScript, null);
    }

    public async Task RemoveContactAsync(RemoveContactRequest request, CancellationToken cancellationToken = default)
    {
        var contactKind = NormalizeContactKind(request.ContactKind) ?? "MailContact";
        var identity = EscapePs(request.Identity);
        var command = contactKind.Equals("MailUser", StringComparison.OrdinalIgnoreCase)
            ? "Remove-MailUser"
            : "Remove-MailContact";

        var script = $@"
{command} -Identity '{identity}' -Confirm:$false -ErrorAction Stop
Write-Output 'OK'";

        await RunScriptAsync(script, cancellationToken);
    }

    private static string? NormalizeContactKind(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        if (string.Equals(value, "MailContact", StringComparison.OrdinalIgnoreCase))
        {
            return "MailContact";
        }

        if (string.Equals(value, "MailUser", StringComparison.OrdinalIgnoreCase))
        {
            return "MailUser";
        }

        return null;
    }

    private static string NormalizeContactSortProperty(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "DisplayName";
        }

        return value.Trim() switch
        {
            "Alias" => "Alias",
            "PrimarySmtpAddress" => "PrimarySmtpAddress",
            "ExternalEmailAddress" => "ExternalEmailAddress",
            "UserPrincipalName" => "UserPrincipalName",
            "ContactKind" => "ContactKind",
            _ => "DisplayName"
        };
    }
}
