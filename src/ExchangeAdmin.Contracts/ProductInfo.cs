using System.Reflection;

namespace ExchangeAdmin.Contracts;

public static class ProductInfo
{
    public static string Version
    {
        get
        {
            var informationalVersion = typeof(ProductInfo).Assembly
                .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
                .InformationalVersion;

            if (!string.IsNullOrWhiteSpace(informationalVersion))
            {
                return informationalVersion.Split('+')[0];
            }

            var assemblyVersion = typeof(ProductInfo).Assembly.GetName().Version;
            return assemblyVersion is null
                ? "0.0.0"
                : $"{assemblyVersion.Major}.{assemblyVersion.Minor}.{Math.Max(assemblyVersion.Build, 0)}";
        }
    }

    public static string DisplayVersion => $"v{Version}";
}
