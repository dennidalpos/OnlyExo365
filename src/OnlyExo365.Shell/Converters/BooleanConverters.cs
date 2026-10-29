using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace OnlyExo365.Shell.Converters;

             
                                
              
public class BoolToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is bool b)
        {
            return b ? Visibility.Visible : Visibility.Collapsed;
        }
        return Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is Visibility v)
        {
            return v == Visibility.Visible;
        }
        return false;
    }
}

             
                                          
              
public class InverseBoolToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is bool b)
        {
            return b ? Visibility.Collapsed : Visibility.Visible;
        }
        return Visibility.Visible;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is Visibility v)
        {
            return v != Visibility.Visible;
        }
        return true;
    }
}

             
                               
              
public class InverseBoolConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is bool b)
        {
            return !b;
        }
        return true;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is bool b)
        {
            return !b;
        }
        return false;
    }
}

/// <summary>Joins strings with commas.</summary>
public class ListToStringConverter : IValueConverter
{
    private static readonly string EmptyResult = "-";

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is not IEnumerable<string> enumerable)
            return EmptyResult;

        // Reuse lists/arrays.
        if (value is IList<string> list)
        {
            if (list.Count == 0)
                return EmptyResult;
            return string.Join(", ", list);
        }

        if (value is string[] array)
        {
            if (array.Length == 0)
                return EmptyResult;
            return string.Join(", ", array);
        }

        using var enumerator = enumerable.GetEnumerator();
        if (!enumerator.MoveNext())
            return EmptyResult;

        return string.Join(", ", enumerable);
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return Binding.DoNothing;
    }
}

/// <summary>Collapses null, empty strings and collections.</summary>
public class NullToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value == null)
            return Visibility.Collapsed;

        if (value is string str)
            return string.IsNullOrEmpty(str) ? Visibility.Collapsed : Visibility.Visible;

        // Avoid iterator allocations.
        if (value is System.Collections.ICollection collection)
            return collection.Count == 0 ? Visibility.Collapsed : Visibility.Visible;

        if (value is System.Collections.IEnumerable enumerable)
        {
            var enumerator = enumerable.GetEnumerator();
            try
            {
                return enumerator.MoveNext() ? Visibility.Visible : Visibility.Collapsed;
            }
            finally
            {
                (enumerator as IDisposable)?.Dispose();
            }
        }

        return Visibility.Visible;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return Binding.DoNothing;
    }
}

             
                                                                    
              
public class ZeroToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is int intValue)
            return intValue == 0 ? Visibility.Collapsed : Visibility.Visible;

        if (value is long longValue)
            return longValue == 0 ? Visibility.Collapsed : Visibility.Visible;

        if (value is double doubleValue)
            return doubleValue == 0 ? Visibility.Collapsed : Visibility.Visible;

        return Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return Binding.DoNothing;
    }
}

