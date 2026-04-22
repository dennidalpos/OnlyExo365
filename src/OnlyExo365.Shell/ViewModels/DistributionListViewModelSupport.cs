using System.Collections.ObjectModel;
using System.Linq;

namespace OnlyExo365.Shell.ViewModels;

internal static class DistributionListViewModelSupport
{
    public static string MapGroupTypeForWorker(string? groupType)
    {
        if (string.Equals(groupType, "Dynamic", StringComparison.OrdinalIgnoreCase))
        {
            return "DynamicDistributionGroup";
        }

        if (string.Equals(groupType, "Microsoft365", StringComparison.OrdinalIgnoreCase))
        {
            return "UnifiedGroup";
        }

        return "DistributionGroup";
    }

    public static string FormatGroupTypeLabel(string? groupType) => groupType switch
    {
        "MailSecurity" => "Mail-enabled security",
        "Microsoft365" => "Microsoft 365",
        "Dynamic" => "Dynamic distribution",
        "Distribution" => "Distribution",
        _ => groupType ?? string.Empty
    };

    public static IEnumerable<string> NormalizeSenderList(IEnumerable<string> list)
    {
        return list
            .Select(value => value?.Trim())
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)!;
    }

    public static bool SenderListEquals(IEnumerable<string> current, IEnumerable<string> original)
        => NormalizeSenderList(current).SequenceEqual(NormalizeSenderList(original), StringComparer.OrdinalIgnoreCase);

    public static void ResetObservableList(ObservableCollection<string> target, IEnumerable<string> values)
    {
        target.Clear();
        foreach (var value in values)
        {
            target.Add(value);
        }
    }
}

