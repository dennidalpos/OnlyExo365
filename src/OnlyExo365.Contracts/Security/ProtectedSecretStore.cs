using System.Security.Cryptography;
using System.Text;
using System.Runtime.Versioning;

namespace OnlyExo365.Contracts.Security;

public sealed class ProtectedSecretReference
{
    public string Id { get; set; } = string.Empty;
}

public static class ProtectedSecretStore
{
    private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("OnlyExo365.OnlyExo365.IPC.Secret.v1");

    public static ProtectedSecretReference? Create(string? secret)
    {
        if (string.IsNullOrWhiteSpace(secret))
        {
            return null;
        }

        if (!IsWindows())
        {
            throw new PlatformNotSupportedException("ProtectedSecretStore requires Windows DPAPI support.");
        }
        CleanupExpiredSecrets();

        var id = Guid.NewGuid().ToString("N");
        var path = GetSecretPath(id);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);

        var clearBytes = Encoding.UTF8.GetBytes(secret);
        var protectedBytes = ProtectedData.Protect(clearBytes, Entropy, DataProtectionScope.CurrentUser);
        File.WriteAllBytes(path, protectedBytes);
        CryptographicOperations.ZeroMemory(clearBytes);

        return new ProtectedSecretReference { Id = id };
    }

    public static string? Consume(ProtectedSecretReference? reference)
    {
        var path = TryGetSecretPath(reference);
        if (path == null)
        {
            return null;
        }

        if (!IsWindows())
        {
            throw new PlatformNotSupportedException("ProtectedSecretStore requires Windows DPAPI support.");
        }

        if (!File.Exists(path))
        {
            return null;
        }

        byte[]? protectedBytes = null;
        byte[]? clearBytes = null;

        try
        {
            protectedBytes = File.ReadAllBytes(path);
            clearBytes = ProtectedData.Unprotect(protectedBytes, Entropy, DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(clearBytes);
        }
        finally
        {
            if (protectedBytes is not null)
            {
                CryptographicOperations.ZeroMemory(protectedBytes);
            }

            if (clearBytes is not null)
            {
                CryptographicOperations.ZeroMemory(clearBytes);
            }

            TryDelete(reference);
        }
    }

    public static void TryDelete(ProtectedSecretReference? reference)
    {
        var path = TryGetSecretPath(reference);
        if (path == null)
        {
            return;
        }

        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
        }
    }

    public static bool Exists(ProtectedSecretReference? reference)
    {
        var path = TryGetSecretPath(reference);
        return path != null && File.Exists(path);
    }

    private static string GetSecretPath(string id)
        => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OnlyExo365",
            "ipc-secrets",
            $"{id}.bin");

    private static string? TryGetSecretPath(ProtectedSecretReference? reference)
    {
        if (reference == null || string.IsNullOrWhiteSpace(reference.Id))
        {
            return null;
        }

        var trimmedId = reference.Id.Trim();
        return Guid.TryParseExact(trimmedId, "N", out var parsedId)
            ? GetSecretPath(parsedId.ToString("N"))
            : null;
    }

    private static void CleanupExpiredSecrets()
    {
        try
        {
            var directory = Path.GetDirectoryName(GetSecretPath("placeholder"));
            if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
            {
                return;
            }

            var thresholdUtc = DateTime.UtcNow.AddHours(-6);
            foreach (var file in Directory.EnumerateFiles(directory, "*.bin"))
            {
                try
                {
                    if (File.GetCreationTimeUtc(file) < thresholdUtc)
                    {
                        File.Delete(file);
                    }
                }
                catch
                {
                }
            }
        }
        catch
        {
        }
    }

    [SupportedOSPlatformGuard("windows")]
    private static bool IsWindows() => OperatingSystem.IsWindows();
}

