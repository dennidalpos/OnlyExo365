using OnlyExo365.Shell.Localization;
using OnlyExo365.Shell.Services;

namespace OnlyExo365.Shell.ViewModels;

public sealed class LanguageSelectionViewModel : ViewModelBase
{
    private readonly UserPreferencesService _preferences;

    public LanguageSelectionViewModel(UserPreferencesService preferences)
    {
        _preferences = preferences;
        AvailableLocales = LocalizationService.Instance.AvailableLocales;
        _selectedLocale = AvailableLocales.First(l => l.Code == LocalizationService.Instance.CurrentLocale);
    }

    public IReadOnlyList<LocaleOption> AvailableLocales { get; }

    private LocaleOption _selectedLocale;
    public LocaleOption SelectedLocale
    {
        get => _selectedLocale;
        set
        {
            if (SetProperty(ref _selectedLocale, value) && value is not null)
            {
                LocalizationService.Instance.SetLocale(value.Code);
                _preferences.SaveLocale(value.Code);
            }
        }
    }
}

