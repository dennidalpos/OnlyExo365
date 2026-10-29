namespace OnlyExo365.Worker.PowerShell;

internal static class ExoRequestSanitizer
{
    private static readonly HashSet<string> AllowedMailboxRecipientTypeDetails = new(StringComparer.OrdinalIgnoreCase)
    {
        "UserMailbox",
        "SharedMailbox",
        "RoomMailbox",
        "EquipmentMailbox"
    };

    private static readonly HashSet<string> AllowedMailboxSortProperties = new(StringComparer.OrdinalIgnoreCase)
    {
        "DisplayName",
        "PrimarySmtpAddress",
        "Alias",
        "UserPrincipalName",
        "RecipientTypeDetails"
    };

    private static readonly HashSet<string> AllowedGroupSortProperties = new(StringComparer.OrdinalIgnoreCase)
    {
        "DisplayName",
        "PrimarySmtpAddress",
        "Alias",
        "RecipientTypeDetails"
    };

    internal static string? NormalizeMailboxRecipientTypeDetails(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var normalized = value.Trim();
        return AllowedMailboxRecipientTypeDetails.Contains(normalized)
            ? AllowedMailboxRecipientTypeDetails.First(candidate => string.Equals(candidate, normalized, StringComparison.OrdinalIgnoreCase))
            : null;
    }

    internal static string NormalizeMailboxSortProperty(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "DisplayName";
        }

        var normalized = value.Trim();
        return AllowedMailboxSortProperties.Contains(normalized)
            ? AllowedMailboxSortProperties.First(candidate => string.Equals(candidate, normalized, StringComparison.OrdinalIgnoreCase))
            : "DisplayName";
    }

    internal static string NormalizeGroupSortProperty(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "DisplayName";
        }

        var normalized = value.Trim();
        return AllowedGroupSortProperties.Contains(normalized)
            ? AllowedGroupSortProperties.First(candidate => string.Equals(candidate, normalized, StringComparison.OrdinalIgnoreCase))
            : "DisplayName";
    }

    internal static string FormatStringArrayParameter(IEnumerable<string> values)
    {
        ArgumentNullException.ThrowIfNull(values);

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
}

