using System.Windows;
using System.Windows.Markup;

namespace ExchangeAdmin.Presentation.Localization;

[MarkupExtensionReturnType(typeof(string))]
public sealed class LocExtension : MarkupExtension
{
    public string Key { get; set; } = string.Empty;

    public override object ProvideValue(IServiceProvider serviceProvider)
    {
        if (serviceProvider.GetService(typeof(IProvideValueTarget)) is IProvideValueTarget pvt
            && pvt.TargetObject is DependencyObject targetObject
            && pvt.TargetProperty is DependencyProperty targetProperty)
        {
            var weakRef = new WeakReference<DependencyObject>(targetObject);
            var key = Key;
            LocalizationService.Instance.CultureChanged += (_, _) =>
            {
                if (weakRef.TryGetTarget(out var t) && t is not null)
                {
                    t.Dispatcher?.Invoke(
                        () => t.SetValue(targetProperty, LocalizationService.Instance.Get(key)));
                }
            };
        }

        return LocalizationService.Instance.Get(Key);
    }
}
