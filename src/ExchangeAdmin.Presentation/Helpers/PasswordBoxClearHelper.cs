using System.Windows;
using System.Windows.Controls;

namespace ExchangeAdmin.Presentation.Helpers;

public static class PasswordBoxClearHelper
{
    public static readonly DependencyProperty ClearTriggerProperty =
        DependencyProperty.RegisterAttached(
            "ClearTrigger",
            typeof(int),
            typeof(PasswordBoxClearHelper),
            new PropertyMetadata(0, OnClearTriggerChanged));

    public static int GetClearTrigger(DependencyObject obj)
    {
        return (int)obj.GetValue(ClearTriggerProperty);
    }

    public static void SetClearTrigger(DependencyObject obj, int value)
    {
        obj.SetValue(ClearTriggerProperty, value);
    }

    private static void OnClearTriggerChanged(DependencyObject dependencyObject, DependencyPropertyChangedEventArgs args)
    {
        if (dependencyObject is PasswordBox passwordBox &&
            args.NewValue is int nextValue &&
            args.OldValue is int previousValue &&
            nextValue != previousValue)
        {
            passwordBox.Clear();
        }
    }
}
