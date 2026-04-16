using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

internal static class OrganizationDirectoryCommandBuilder
{
    public static string BuildUpsertAddressListScript(UpsertAddressListRequest request)
    {
        var identity = EscapePs(request.Identity);
        var name = EscapePs(request.Name);
        var displayName = EscapePs(request.DisplayName);
        var recipientFilter = EscapePs(request.RecipientFilter);
        var recipientContainer = EscapePs(request.RecipientContainer);
        var includedRecipients = ToPsArrayLiteral(request.IncludedRecipients);
        var conditionalCompany = ToPsArrayLiteral(request.ConditionalCompany);
        var conditionalDepartment = ToPsArrayLiteral(request.ConditionalDepartment);
        var conditionalStateOrProvince = ToPsArrayLiteral(request.ConditionalStateOrProvince);
        var conditionalCustomAttribute1 = ToPsArrayLiteral(request.ConditionalCustomAttribute1);

        return $@"
$includedRecipients = {includedRecipients}
$conditionalCompany = {conditionalCompany}
$conditionalDepartment = {conditionalDepartment}
$conditionalStateOrProvince = {conditionalStateOrProvince}
$conditionalCustomAttribute1 = {conditionalCustomAttribute1}
$params = @{{}}
if ('{displayName}' -ne '') {{
    $params['DisplayName'] = '{displayName}'
}}
if ('{recipientContainer}' -ne '') {{
    $params['RecipientContainer'] = '{recipientContainer}'
}}
if ('{recipientFilter}' -ne '') {{
    $params['RecipientFilter'] = '{recipientFilter}'
}} else {{
    if ($includedRecipients.Count -gt 0) {{
        $params['IncludedRecipients'] = $includedRecipients
    }}
    if ($conditionalCompany.Count -gt 0) {{
        $params['ConditionalCompany'] = $conditionalCompany
    }}
    if ($conditionalDepartment.Count -gt 0) {{
        $params['ConditionalDepartment'] = $conditionalDepartment
    }}
    if ($conditionalStateOrProvince.Count -gt 0) {{
        $params['ConditionalStateOrProvince'] = $conditionalStateOrProvince
    }}
    if ($conditionalCustomAttribute1.Count -gt 0) {{
        $params['ConditionalCustomAttribute1'] = $conditionalCustomAttribute1
    }}
}}
if ('{identity}' -ne '') {{
    Set-AddressList -Identity '{identity}' @params -ErrorAction Stop
}} else {{
    $params['Name'] = '{name}'
    New-AddressList @params -ErrorAction Stop
}}";
    }

    public static string BuildUpsertAddressBookPolicyScript(UpsertAddressBookPolicyRequest request)
    {
        var identity = EscapePs(request.Identity);
        var name = EscapePs(request.Name);
        var globalAddressList = EscapePs(request.GlobalAddressList);
        var offlineAddressBook = EscapePs(request.OfflineAddressBook);
        var roomList = EscapePs(request.RoomList);
        var addressLists = ToPsArrayLiteral(request.AddressLists);

        return $@"
$addressLists = {addressLists}
$params = @{{
    AddressLists = $addressLists
    GlobalAddressList = '{globalAddressList}'
    OfflineAddressBook = '{offlineAddressBook}'
}}
if ('{roomList}' -ne '') {{
    $params['RoomList'] = '{roomList}'
}}
if ('{identity}' -ne '') {{
    Set-AddressBookPolicy -Identity '{identity}' @params -ErrorAction Stop
}} else {{
    $params['Name'] = '{name}'
    New-AddressBookPolicy @params -ErrorAction Stop
}}";
    }

    public static string BuildUpsertOfflineAddressBookScript(UpsertOfflineAddressBookRequest request)
    {
        var identity = EscapePs(request.Identity);
        var name = EscapePs(request.Name);
        var addressLists = ToPsArrayLiteral(request.AddressLists);
        var diffRetentionPeriod = ToPsNullableIntLiteral(request.DiffRetentionPeriod);

        return $@"
$addressLists = {addressLists}
$params = @{{
    AddressLists = $addressLists
}}
if ($null -ne {diffRetentionPeriod}) {{
    $params['DiffRetentionPeriod'] = {diffRetentionPeriod}
}}
if ('{identity}' -ne '') {{
    Set-OfflineAddressBook -Identity '{identity}' @params -ErrorAction Stop
}} else {{
    $params['Name'] = '{name}'
    New-OfflineAddressBook @params -ErrorAction Stop
}}";
    }

    public static string BuildUpsertSharingPolicyScript(UpsertSharingPolicyRequest request)
    {
        var identity = EscapePs(request.Identity);
        var name = EscapePs(request.Name);
        var domains = ToPsArrayLiteral(request.Domains);
        var enabled = ToPsBoolLiteral(request.Enabled);
        var makeDefault = ToPsBoolLiteral(request.MakeDefault);

        return $@"
$domains = {domains}
$params = @{{
    Domains = $domains
    Enabled = {enabled}
}}
if ('{identity}' -ne '') {{
    Set-SharingPolicy -Identity '{identity}' @params -ErrorAction Stop
    if ({makeDefault}) {{
        Set-SharingPolicy -Identity '{identity}' -Default:$true -ErrorAction Stop
    }}
}} else {{
    $params['Name'] = '{name}'
    if ({makeDefault}) {{
        $params['Default'] = $true
    }}
    New-SharingPolicy @params -ErrorAction Stop
}}";
    }

    private static string EscapePs(string? value)
        => (value ?? string.Empty).Replace("'", "''");

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

    private static string ToPsBoolLiteral(bool value)
        => value ? "$true" : "$false";

    private static string ToPsNullableIntLiteral(int? value)
        => value.HasValue ? value.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) : "$null";
}
