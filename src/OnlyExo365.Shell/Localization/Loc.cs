namespace OnlyExo365.Shell.Localization;

public static class Loc
{
    public static string Get(string key) => LocalizationService.Instance.Get(key);
    public static string GetFormat(string key, params object[] args) => LocalizationService.Instance.GetFormat(key, args);
}

