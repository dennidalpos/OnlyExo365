using System.ComponentModel;
using System.Windows;
using System.Windows.Media;

namespace OnlyExo365.Shell.Services;

internal sealed class ThemePaletteService : IDisposable
{
    private readonly ResourceDictionary _resources;
    private readonly Dictionary<object, Brush> _defaultBrushes;

    public ThemePaletteService(ResourceDictionary resources)
    {
        _resources = resources;
        var theme = resources.MergedDictionaries.FirstOrDefault() ?? resources;
        _defaultBrushes = theme.Keys.Cast<object>()
            .Where(key => theme[key] is Brush)
            .ToDictionary(key => key, key => (Brush)theme[key]);
        Apply(SystemParameters.HighContrast);
        SystemParameters.StaticPropertyChanged += OnSystemPropertyChanged;
    }

    internal void Apply(bool highContrast)
    {
        foreach (var (key, brush) in _defaultBrushes)
            _resources[key] = highContrast ? GetSystemBrush(key.ToString()!) : brush;
    }

    private static Brush GetSystemBrush(string key) => key switch
    {
        "AccentBrush" or "AccentHoverBrush" or "PrimaryBrush" or
            "DangerBackgroundBrush" or "DangerHoverBackgroundBrush" => SystemColors.HighlightBrush,
        "PrimaryActionTextBrush" or "DangerActionTextBrush" => SystemColors.HighlightTextBrush,
        "BorderBrush" or "PrimaryTextBrush" or "SecondaryTextBrush" or
            "SuccessBrush" or "WarningBrush" or "ErrorBrush" or "InfoBrush" => SystemColors.WindowTextBrush,
        _ => SystemColors.WindowBrush
    };

    private void OnSystemPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(SystemParameters.HighContrast))
            Application.Current?.Dispatcher.Invoke(() => Apply(SystemParameters.HighContrast));
    }

    public void Dispose() => SystemParameters.StaticPropertyChanged -= OnSystemPropertyChanged;
}
