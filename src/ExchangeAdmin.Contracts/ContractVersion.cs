namespace ExchangeAdmin.Contracts;

/// <summary>
/// Versione del contratto IPC. Deve corrispondere tra UI e Worker.
/// </summary>
public static class ContractVersion
{
    public const int Major = 1;
    public const int Minor = 0;
    public const int Patch = 0;

    public static string Version => $"{Major}.{Minor}.{Patch}";

    public static bool IsCompatible(string otherVersion)
    {
        if (string.IsNullOrWhiteSpace(otherVersion))
            return false;

        var parts = otherVersion.Split('.');
        if (parts.Length < 2)
            return false;

        if (!int.TryParse(parts[0], out var otherMajor))
            return false;

        // Compatibilità: stesso major version
        return otherMajor == Major;
    }
}
