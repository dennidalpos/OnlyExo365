using System.Globalization;
using System.Resources;

namespace OnlyExo365.Shell.Localization;

public sealed class LocalizationService : ILocalizationService
{
    private static readonly ResourceManager _rm =
        new("OnlyExo365.Shell.Localization.Strings",
            typeof(LocalizationService).Assembly);

    private string _currentLocale = "en";

    public static LocalizationService Instance { get; } = new();

    private LocalizationService() { }

    public string CurrentLocale => _currentLocale;

    public IReadOnlyList<LocaleOption> AvailableLocales { get; } =
    [
        new LocaleOption("en", "English"),
        new LocaleOption("it", "Italiano")
    ];

    public string Get(string key)
    {
        var culture = CultureInfo.GetCultureInfo(_currentLocale);
        return _rm.GetString(key, culture) ?? key;
    }

    public string GetFormat(string key, params object[] args)
    {
        var template = Get(key);
        return args.Length == 0 ? template : string.Format(template, args);
    }

    public void SetLocale(string localeCode)
    {
        if (_currentLocale == localeCode) return;
        _currentLocale = localeCode;
        CultureChanged?.Invoke(this, EventArgs.Empty);
    }

    public event EventHandler? CultureChanged;
}

