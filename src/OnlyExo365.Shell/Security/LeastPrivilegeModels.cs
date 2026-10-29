using OnlyExo365.Contracts;

namespace OnlyExo365.Shell.Security;

public enum LeastPrivilegeFeatureStatus
{
    PendingSession,
    Available,
    NeedsAdditionalSession,
    Blocked
}

public sealed class LeastPrivilegeFeatureDefinition
{
    public required string FeatureId { get; init; }

    public required string ModuleName { get; init; }

    public required string FeatureName { get; init; }

    public required string Description { get; init; }

    public IReadOnlyList<ExchangeAuthenticationMode> AllowedAuthenticationModes { get; init; } =
        Array.Empty<ExchangeAuthenticationMode>();

    public IReadOnlyList<string> RequiredCmdletsAll { get; init; } = Array.Empty<string>();

    public IReadOnlyList<string> RequiredCmdletsAny { get; init; } = Array.Empty<string>();

    public IReadOnlyList<string> RecommendedExchangeRoles { get; init; } = Array.Empty<string>();

    public IReadOnlyList<string> RequiredGraphScopes { get; init; } = Array.Empty<string>();

    public IReadOnlyList<string> RecommendedPurviewRoles { get; init; } = Array.Empty<string>();

    public IReadOnlyList<string> RecommendedDefenderRoles { get; init; } = Array.Empty<string>();

    public IReadOnlyList<string> Dependencies { get; init; } = Array.Empty<string>();

    public string? Notes { get; init; }

    public bool RequiresAdditionalSessionValidation { get; init; }
}

public sealed class LeastPrivilegeFeatureEvaluation
{
    public LeastPrivilegeFeatureEvaluation(
        LeastPrivilegeFeatureDefinition definition,
        LeastPrivilegeFeatureStatus status,
        IReadOnlyList<string> missingRequirements,
        string validationMessage)
    {
        Definition = definition;
        Status = status;
        MissingRequirements = missingRequirements;
        ValidationMessage = validationMessage;
    }

    public LeastPrivilegeFeatureDefinition Definition { get; }

    public LeastPrivilegeFeatureStatus Status { get; }

    public IReadOnlyList<string> MissingRequirements { get; }

    public string ValidationMessage { get; }

    public string FeatureId => Definition.FeatureId;

    public string ModuleName => Definition.ModuleName;

    public string FeatureName => Definition.FeatureName;

    public string Description => Definition.Description;

    public string StatusLabel => Status switch
    {
        LeastPrivilegeFeatureStatus.Available => "Ready",
        LeastPrivilegeFeatureStatus.NeedsAdditionalSession => "Needs session",
        LeastPrivilegeFeatureStatus.Blocked => "Blocked",
        _ => "Pending"
    };

    public string StatusColor => Status switch
    {
        LeastPrivilegeFeatureStatus.Available => "#4EC9B0",
        LeastPrivilegeFeatureStatus.NeedsAdditionalSession => "#DCDCAA",
        LeastPrivilegeFeatureStatus.Blocked => "#F14C4C",
        _ => "#9D9D9D"
    };

    public bool IsNavigationAllowed => Status != LeastPrivilegeFeatureStatus.Blocked;

    public bool HasMissingRequirements => MissingRequirements.Count > 0;

    public string MissingRequirementsDisplay => MissingRequirements.Count == 0
        ? "None"
        : string.Join("; ", MissingRequirements);

    public string AllowedAuthenticationModesDisplay => Definition.AllowedAuthenticationModes.Count == 0
        ? "Current auth modes"
        : string.Join(", ", Definition.AllowedAuthenticationModes);

    public string RequiredCmdletsDisplay => FormatRequirementList(
        Definition.RequiredCmdletsAll,
        Definition.RequiredCmdletsAny);

    public string ExchangeRolesDisplay => FormatList(Definition.RecommendedExchangeRoles);

    public string GraphScopesDisplay => FormatList(Definition.RequiredGraphScopes);

    public string PurviewRolesDisplay => FormatList(Definition.RecommendedPurviewRoles);

    public string DefenderRolesDisplay => FormatList(Definition.RecommendedDefenderRoles);

    public string DependenciesDisplay => FormatList(Definition.Dependencies);

    public string NotesDisplay => string.IsNullOrWhiteSpace(Definition.Notes)
        ? "Exchange RBAC role membership is advisory and must be validated in the tenant."
        : Definition.Notes!;

    private static string FormatRequirementList(
        IReadOnlyList<string> requiredAll,
        IReadOnlyList<string> requiredAny)
    {
        if (requiredAll.Count == 0 && requiredAny.Count == 0)
        {
            return "None";
        }

        var parts = new List<string>();
        if (requiredAll.Count > 0)
        {
            parts.Add($"All: {string.Join(", ", requiredAll)}");
        }

        if (requiredAny.Count > 0)
        {
            parts.Add($"Any: {string.Join(", ", requiredAny)}");
        }

        return string.Join(" | ", parts);
    }

    private static string FormatList(IReadOnlyList<string> items)
        => items.Count == 0 ? "None" : string.Join(", ", items);
}

