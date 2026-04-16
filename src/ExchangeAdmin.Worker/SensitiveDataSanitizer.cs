using System.Text.RegularExpressions;

namespace ExchangeAdmin.Worker;

internal static partial class SensitiveDataSanitizer
{
    private const string RedactedSecret = "<redacted>";

    public static string Sanitize(string message)
    {
        if (string.IsNullOrEmpty(message))
        {
            return message;
        }

        var sanitized = ConvertToSecureStringPattern().Replace(
            message,
            static match => $"{match.Groups["prefix"].Value}'{RedactedSecret}'{match.Groups["suffix"].Value}");

        sanitized = PlainTextPasswordAssignmentPattern().Replace(
            sanitized,
            static match => $"{match.Groups["prefix"].Value}'{RedactedSecret}'");

        return sanitized;
    }

    [GeneratedRegex("(?<prefix>ConvertTo-SecureString\\s+)'[^']*'(?<suffix>\\s+-AsPlainText\\s+-Force)", RegexOptions.IgnoreCase)]
    private static partial Regex ConvertToSecureStringPattern();

    [GeneratedRegex("(?<prefix>-Password\\s+)'[^']*'", RegexOptions.IgnoreCase)]
    private static partial Regex PlainTextPasswordAssignmentPattern();
}
