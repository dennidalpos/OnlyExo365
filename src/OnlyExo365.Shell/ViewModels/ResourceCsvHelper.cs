namespace OnlyExo365.Shell.ViewModels;

internal static class ResourceCsvHelper
{
    public static List<string> Parse(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return [];
        }

        return value
            .Split([',', ';', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(static entry => !string.IsNullOrWhiteSpace(entry))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public static HashSet<string> ToSet(string? value)
        => Parse(value).ToHashSet(StringComparer.OrdinalIgnoreCase);

    public static string ToCsv(IEnumerable<string>? values)
    {
        if (values == null)
        {
            return string.Empty;
        }

        return string.Join(
            ", ",
            values
                .Where(static value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase));
    }

    public static string NormalizeCsv(string? value)
    {
        return string.Join(
            "|",
            Parse(value)
                .OrderBy(static item => item, StringComparer.OrdinalIgnoreCase));
    }
}

