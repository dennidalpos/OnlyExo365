using System.IO;

namespace ExchangeAdmin.Presentation.Services;

public enum CatalogAutoUpdateMode
{
    Disabled,
    Daily,
    Monthly
}

public sealed class LicenseCatalogConfiguration
{
    public const string SectionName = "licensingCatalog";

    public const string DefaultRemoteSource =
        "https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference";

    public CatalogAutoUpdateMode AutoUpdateMode { get; set; } = CatalogAutoUpdateMode.Daily;

    public bool CheckOnStartup { get; set; } = true;

    public string RemoteSource { get; set; } = DefaultRemoteSource;

    public int DownloadTimeoutSeconds { get; set; } = 30;

    public string? LocalCachePath { get; set; }

    public string ResolveLocalCachePath() =>
        string.IsNullOrWhiteSpace(LocalCachePath)
            ? Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OnlyExo365",
                "LicenseCatalog")
            : LocalCachePath;

    public static LicenseCatalogConfiguration CreateDefault() => new();
}
