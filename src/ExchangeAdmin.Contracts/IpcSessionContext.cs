using System.Diagnostics;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;

namespace ExchangeAdmin.Contracts;

public sealed class IpcSessionContext
{
    public required string UserScope { get; init; }

    public required int SessionId { get; init; }

    public string RequestPipeName => BuildPipeName(IpcConstants.PipeName);

    public string EventPipeName => BuildPipeName(IpcConstants.EventPipeName);

    public static IpcSessionContext CreateForCurrentProcess()
    {
        return new IpcSessionContext
        {
            UserScope = ResolveUserScope(),
            SessionId = ResolveSessionId()
        };
    }

    public bool Matches(IpcSessionContext? other)
    {
        return other != null &&
               SessionId == other.SessionId &&
               string.Equals(UserScope, other.UserScope, StringComparison.Ordinal);
    }

    private string BuildPipeName(string basePipeName)
    {
        var scopeHash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(UserScope)))[..16];
        return $"{basePipeName}_{scopeHash}_{SessionId}";
    }

    private static string ResolveUserScope()
    {
        try
        {
            if (OperatingSystem.IsWindows())
            {
                return WindowsIdentity.GetCurrent().User?.Value ?? Environment.UserName;
            }
        }
        catch
        {
        }

        return Environment.UserName;
    }

    private static int ResolveSessionId()
    {
        try
        {
            return Process.GetCurrentProcess().SessionId;
        }
        catch
        {
            return 0;
        }
    }
}
