using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Errors;
using OnlyExo365.Shell.Text;
using OnlyExo365.Shell.ViewModels;

namespace OnlyExo365.Shell.Services;

             
                                                       
              
public class ErrorDialogService
{
    internal static Func<string, string, bool>? ConfirmationHandlerOverride { get; set; }
    internal static Action<string, string>? WarningHandlerOverride { get; set; }
    internal static Action<string, string>? InfoHandlerOverride { get; set; }

                 
                                                         
                  
    public static void ShowError(string title, string message, string? details = null)
    {
        ShowNonBlockingMessageWindow(UserMessageCatalog.CombineMessageAndDetails(message, details), title, MessageBoxImage.Error);
    }

                 
                                                        
                  
    public static void ShowError(string title, NormalizedErrorDto error)
    {
        var message = GetUserFriendlyMessage(error);
        var details = error.Details;

        ShowError(title, message, details);
    }

                 
                                                     
                  
    public static void ShowError(string title, NormalizedError error)
    {
        var message = GetUserFriendlyMessage(error.ToDto());
        var details = error.Details;

        ShowError(title, message, details);
    }

                 
                               
                  
    public static void ShowWarning(string title, string message)
    {
        if (WarningHandlerOverride != null)
        {
            WarningHandlerOverride(title, message);
            return;
        }

        if (TryGetShellPrompt(out var prompt))
        {
            prompt.ShowWarning(title, message);
            return;
        }

        ShowNonBlockingMessageWindow(message, title, MessageBoxImage.Warning);
    }

                 
                                    
                  
    public static void ShowInfo(string title, string message)
    {
        if (InfoHandlerOverride != null)
        {
            InfoHandlerOverride(title, message);
            return;
        }

        if (TryGetShellPrompt(out var prompt))
        {
            prompt.ShowInformation(title, message);
            return;
        }

        ShowNonBlockingMessageWindow(message, title, MessageBoxImage.Information);
    }

                 
                                                              
                  
    public static bool ShowConfirmation(string title, string message)
    {
        if (ConfirmationHandlerOverride != null)
        {
            return ConfirmationHandlerOverride(title, message);
        }

        if (TryGetShellPrompt(out var prompt))
        {
            return prompt.ShowConfirmationBlocking(title, message);
        }

        var result = MessageBox.Show(
            message,
            title,
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        return result == MessageBoxResult.Yes;
    }

    private static bool TryGetShellPrompt(out ShellPromptViewModel prompt)
    {
        prompt = null!;

        if (System.Windows.Application.Current?.MainWindow?.DataContext is not ShellViewModel shell)
        {
            return false;
        }

        prompt = shell.Prompt;
        return true;
    }

    private static void ShowNonBlockingMessageWindow(string message, string title, MessageBoxImage icon)
    {
        void ShowWindow()
        {
            var owner = System.Windows.Application.Current?.MainWindow;
            var window = new Window
            {
                Title = title,
                Owner = owner,
                WindowStartupLocation = owner != null
                    ? WindowStartupLocation.CenterOwner
                    : WindowStartupLocation.CenterScreen,
                SizeToContent = SizeToContent.WidthAndHeight,
                ResizeMode = ResizeMode.NoResize,
                WindowStyle = WindowStyle.ToolWindow,
                ShowInTaskbar = false,
                Topmost = true
            };

            var panel = new StackPanel
            {
                Margin = new Thickness(16),
                Orientation = Orientation.Vertical
            };

            var text = new TextBlock
            {
                Text = message,
                TextWrapping = TextWrapping.Wrap,
                MaxWidth = 420
            };

            var button = new Button
            {
                Content = UiTextCatalog.OkButton,
                Width = 80,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 12, 0, 0),
                IsDefault = true
            };
            button.Click += (_, _) => window.Close();

            panel.Children.Add(text);
            panel.Children.Add(button);
            window.Content = panel;

            window.Show();
        }

        var dispatcher = System.Windows.Application.Current?.Dispatcher;
        if (dispatcher != null && !dispatcher.HasShutdownStarted)
        {
            dispatcher.BeginInvoke((Action)ShowWindow, DispatcherPriority.Normal);
        }
        else
        {
            ShowWindow();
        }
    }

                 
                                                                    
                  
    private static string GetUserFriendlyMessage(NormalizedErrorDto error)
    {
        return UserMessageCatalog.GetFriendlyErrorMessage(error);
    }
}

