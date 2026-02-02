using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Errors;

namespace ExchangeAdmin.Presentation.Services;

             
                                                       
              
public class ErrorDialogService
{
                 
                                                         
                  
    public static void ShowError(string title, string message, string? details = null)
    {
        var displayMessage = message;
        if (!string.IsNullOrEmpty(details))
        {
            displayMessage += $"\n\nDetails:\n{details}";
        }

        ShowNonBlockingMessageWindow(displayMessage, title, MessageBoxImage.Error);
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
        ShowNonBlockingMessageWindow(message, title, MessageBoxImage.Warning);
    }

                 
                                    
                  
    public static void ShowInfo(string title, string message)
    {
        ShowNonBlockingMessageWindow(message, title, MessageBoxImage.Information);
    }

                 
                                                              
                  
    public static bool ShowConfirmation(string title, string message)
    {
        var result = MessageBox.Show(
            message,
            title,
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        return result == MessageBoxResult.Yes;
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
                Content = "OK",
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
                                               
        return error.Code switch
        {
                                    
            ErrorCode.AuthenticationFailed => "Authentication failed. Please check your credentials and try again.",

            ErrorCode.ConditionalAccessBlocked => "Access blocked by Conditional Access policy.\n\n" +
                                                 "Please contact your IT administrator for assistance.",

            ErrorCode.MfaRequired => "Multi-Factor Authentication is required.\n\n" +
                                   "Please complete the MFA verification in the authentication prompt.",

            ErrorCode.TokenExpired => "Your session has expired.\n\n" +
                                    "Please disconnect and reconnect to Exchange Online.",

                                
            ErrorCode.PermissionDenied or ErrorCode.InsufficientPrivileges =>
                "You don't have permission to perform this operation.\n\n" +
                "Please contact your Exchange administrator to request the necessary permissions.",

                               
            ErrorCode.CmdletNotAvailable => "The required PowerShell command is not available.\n\n" +
                                          "Please ensure the ExchangeOnlineManagement module is properly installed.",

            ErrorCode.ModuleNotLoaded => "The ExchangeOnlineManagement module is not loaded.\n\n" +
                                       "Try restarting the worker or reinstalling the module.",

            ErrorCode.InvalidParameter => $"Invalid parameter in the operation.\n\n{error.Message}",

            ErrorCode.OperationNotSupported => "This operation is not supported.\n\n" +
                                             "The feature may not be available with your Exchange configuration.",

                               
            ErrorCode.Throttling => "The Exchange server is limiting the number of requests.\n\n" +
                                  "Please wait a moment. The operation will be retried automatically." +
                                  (error.RetryAfterSeconds.HasValue ? $"\n\nRetry after: {error.RetryAfterSeconds} seconds" : ""),

            ErrorCode.ServiceUnavailable => "The Exchange Online service is temporarily unavailable.\n\n" +
                                          "Please try again in a few moments.",

            ErrorCode.NetworkError => "A network error occurred.\n\n" +
                                    "Please check your internet connection and try again.",

            ErrorCode.Timeout => "The operation timed out.\n\n" +
                               "This may be due to network issues or server load. Please try again.",

                              
            ErrorCode.ResourceNotFound => $"The requested resource was not found.\n\n{error.Message}",

            ErrorCode.ResourceAlreadyExists => $"The resource already exists.\n\n{error.Message}",

                            
            ErrorCode.WorkerNotRunning => "The background worker is not running.\n\n" +
                                        "Please start the worker before connecting to Exchange Online.",

            ErrorCode.WorkerCrashed => "The background worker has stopped unexpectedly.\n\n" +
                                     "Try restarting the worker. Check the logs for more details.",

            ErrorCode.IpcError => "Communication error with the background worker.\n\n" +
                                "Try restarting the worker or the application.",

                      
            ErrorCode.Unknown or _ => !string.IsNullOrEmpty(error.Message)
                                    ? $"An error occurred:\n\n{error.Message}"
                                    : "An unexpected error occurred. Please check the logs for more information."
        };
    }
}
