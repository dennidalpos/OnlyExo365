using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace OnlyExo365.Shell.Localization;

public sealed class LocalizedFormatConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is null || value == DependencyProperty.UnsetValue) return string.Empty;
        return string.Format(culture, Loc.Get((string)parameter), value);
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => Binding.DoNothing;
}
