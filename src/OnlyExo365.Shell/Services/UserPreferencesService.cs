using System.IO;
using System.Text.Json;

namespace OnlyExo365.Shell.Services;

public sealed class UserPreferencesService
{
    private readonly string _filePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "OnlyExo365", "OnlyExo365", "preferences.json");

    public string? LoadLocale()
    {
        try
        {
            if (!File.Exists(_filePath)) return null;
            var json = File.ReadAllText(_filePath);
            using var doc = JsonDocument.Parse(json);
            if (doc.RootElement.TryGetProperty("language", out var el))
                return el.GetString();
            return null;
        }
        catch { return null; }
    }

    public void SaveLocale(string localeCode)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_filePath)!);
            var json = JsonSerializer.Serialize(new { language = localeCode });
            File.WriteAllText(_filePath, json);
        }
        catch { /* non-fatal */ }
    }
}

