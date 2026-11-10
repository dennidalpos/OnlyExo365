namespace OnlyExo365.Shell.Localization;

public interface ILocalizationService
{
    string CurrentLocale { get; }
    string Get(string key);
    string GetFormat(string key, params object[] args);
    IReadOnlyList<LocaleOption> AvailableLocales { get; }
    void SetLocale(string localeCode);
    event EventHandler CultureChanged;
}

public record LocaleOption(string Code, string DisplayName);

