namespace OnlyExo365.Shell.Services;

public sealed class CatalogUpdatedEventArgs : EventArgs
{
    public bool IsSuccess { get; init; }
    public string? Error { get; init; }
    public string? CatalogVersion { get; init; }
    public int EntryCount { get; init; }
    public DateTime? LastUpdatedUtc { get; init; }
    public DateTime? LastCheckedUtc { get; init; }
}

