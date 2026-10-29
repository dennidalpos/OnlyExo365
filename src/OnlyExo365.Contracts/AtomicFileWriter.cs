using System.IO;

namespace OnlyExo365.Contracts;

public static class AtomicFileWriter
{
    public static void Write(string destinationPath, Action<string> writeTemporaryFile)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationPath);
        ArgumentNullException.ThrowIfNull(writeTemporaryFile);

        var fullPath = Path.GetFullPath(destinationPath);
        var directory = Path.GetDirectoryName(fullPath)!;
        Directory.CreateDirectory(directory);
        var temporaryPath = Path.Combine(directory, $".{Path.GetFileName(fullPath)}.{Guid.NewGuid():N}.tmp");

        try
        {
            writeTemporaryFile(temporaryPath);
            if (!File.Exists(temporaryPath))
            {
                throw new IOException("Export did not produce a temporary file.");
            }

            if (File.Exists(fullPath))
            {
                File.Replace(temporaryPath, fullPath, null);
            }
            else
            {
                File.Move(temporaryPath, fullPath);
            }
        }
        finally
        {
            if (File.Exists(temporaryPath))
            {
                File.Delete(temporaryPath);
            }
        }
    }
}
