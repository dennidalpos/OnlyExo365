namespace ExchangeAdmin.Tests;

internal static class TestPathHelper
{
    private static readonly Lazy<string> RepositoryRoot = new(ResolveRepositoryRoot);

    public static string GetRepositoryPath(params string[] segments)
    {
        return Path.GetFullPath(Path.Combine(RepositoryRoot.Value, Path.Combine(segments)));
    }

    private static string ResolveRepositoryRoot()
    {
        foreach (var candidateRoot in EnumerateCandidateRoots())
        {
            if (File.Exists(Path.Combine(candidateRoot, "ExchangeAdmin.sln")) &&
                File.Exists(Path.Combine(candidateRoot, "global.json")) &&
                File.Exists(Path.Combine(candidateRoot, "Directory.Build.props")))
            {
                return candidateRoot;
            }
        }

        throw new DirectoryNotFoundException(
            $"Unable to resolve the repository root from '{AppContext.BaseDirectory}' or '{Directory.GetCurrentDirectory()}'.");
    }

    private static IEnumerable<string> EnumerateCandidateRoots()
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var startPath in new[] { AppContext.BaseDirectory, Directory.GetCurrentDirectory() })
        {
            if (string.IsNullOrWhiteSpace(startPath))
            {
                continue;
            }

            var current = new DirectoryInfo(Path.GetFullPath(startPath));
            while (current is not null)
            {
                if (seen.Add(current.FullName))
                {
                    yield return current.FullName;
                }

                current = current.Parent;
            }
        }
    }
}
