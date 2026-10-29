using System.Windows;
using System.Windows.Markup;

namespace OnlyExo365.Shell.Localization;

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
            EventHandler? handler = null;
            handler = (_, _) =>
            {
                if (!weakRef.TryGetTarget(out var target) || target.Dispatcher.HasShutdownStarted)
                {
                    LocalizationService.Instance.CultureChanged -= handler;
                    return;
                }
                void Update() => target.SetValue(targetProperty, LocalizationService.Instance.Get(key));
                if (target.Dispatcher.CheckAccess()) Update();
                else target.Dispatcher.BeginInvoke(Update);
            };
            LocalizationService.Instance.CultureChanged += handler;
        }

        return LocalizationService.Instance.Get(Key);
    }
}

