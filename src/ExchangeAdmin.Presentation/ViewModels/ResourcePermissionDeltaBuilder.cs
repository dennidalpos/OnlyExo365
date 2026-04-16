using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Presentation.ViewModels;

public static class ResourcePermissionDeltaBuilder
{
    public static HashSet<string> GetOriginalFullAccessUsers(MailboxPermissionsDto? permissions)
    {
        return permissions?.FullAccessPermissions
            .Select(entry => entry.User)
            .Where(static value => !string.IsNullOrWhiteSpace(value))
            .ToHashSet(StringComparer.OrdinalIgnoreCase) ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    }

    public static HashSet<string> GetOriginalSendAsUsers(MailboxPermissionsDto? permissions)
    {
        return permissions?.SendAsPermissions
            .Select(entry => string.IsNullOrWhiteSpace(entry.ResolvedTrustee) ? entry.Trustee : entry.ResolvedTrustee)
            .Where(static value => !string.IsNullOrWhiteSpace(value))
            .ToHashSet(StringComparer.OrdinalIgnoreCase) ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    }

    public static HashSet<string> GetOriginalSendOnBehalfUsers(MailboxPermissionsDto? permissions)
    {
        return permissions?.SendOnBehalfPermissions
            .Select(entry => entry.Identity)
            .Where(static value => !string.IsNullOrWhiteSpace(value))
            .ToHashSet(StringComparer.OrdinalIgnoreCase) ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    }

    public static List<PermissionDeltaActionDto> Build(
        string? fullAccessUsers,
        IEnumerable<string>? originalFullAccessUsers,
        string? sendAsUsers,
        IEnumerable<string>? originalSendAsUsers,
        string? sendOnBehalfUsers,
        IEnumerable<string>? originalSendOnBehalfUsers)
    {
        var actions = new List<PermissionDeltaActionDto>();

        AppendPermissionDelta(
            actions,
            ResourceCsvHelper.ToSet(fullAccessUsers),
            originalFullAccessUsers?.ToHashSet(StringComparer.OrdinalIgnoreCase) ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase),
            PermissionType.FullAccess);

        AppendPermissionDelta(
            actions,
            ResourceCsvHelper.ToSet(sendAsUsers),
            originalSendAsUsers?.ToHashSet(StringComparer.OrdinalIgnoreCase) ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase),
            PermissionType.SendAs);

        AppendPermissionDelta(
            actions,
            ResourceCsvHelper.ToSet(sendOnBehalfUsers),
            originalSendOnBehalfUsers?.ToHashSet(StringComparer.OrdinalIgnoreCase) ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase),
            PermissionType.SendOnBehalf);

        return actions;
    }

    private static void AppendPermissionDelta(
        List<PermissionDeltaActionDto> actions,
        HashSet<string> current,
        HashSet<string> original,
        PermissionType permissionType)
    {
        foreach (var user in current.Except(original, StringComparer.OrdinalIgnoreCase))
        {
            actions.Add(new PermissionDeltaActionDto
            {
                Action = PermissionAction.Add,
                PermissionType = permissionType,
                User = user,
                Description = $"Add {permissionType} -> {user}",
                AutoMapping = permissionType == PermissionType.FullAccess ? true : null
            });
        }

        foreach (var user in original.Except(current, StringComparer.OrdinalIgnoreCase))
        {
            actions.Add(new PermissionDeltaActionDto
            {
                Action = PermissionAction.Remove,
                PermissionType = permissionType,
                User = user,
                Description = $"Remove {permissionType} -> {user}"
            });
        }
    }
}
