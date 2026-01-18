using System.Windows;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Errors;

namespace ExchangeAdmin.Presentation.Services;

/// <summary>
/// Service for displaying user-friendly error dialogs.
/// </summary>
public class ErrorDialogService
{
    /// <summary>
    /// Shows an error dialog with user-friendly message.
    /// </summary>
    public static void ShowError(string title, string message, string? details = null)
    {
        var displayMessage = message;
        if (!string.IsNullOrEmpty(details))
        {
            displayMessage += $"\n\nDetails:\n{details}";
        }

        MessageBox.Show(
            displayMessage,
            title,
            MessageBoxButton.OK,
            MessageBoxImage.Error);
    }

    /// <summary>
    /// Shows an error dialog from a NormalizedErrorDto.
    /// </summary>
    public static void ShowError(string title, NormalizedErrorDto error)
    {
        var message = GetUserFriendlyMessage(error);
        var details = error.Details;

        ShowError(title, message, details);
    }

    /// <summary>
    /// Shows an error dialog from a NormalizedError.
    /// </summary>
    public static void ShowError(string title, NormalizedError error)
    {
        var message = GetUserFriendlyMessage(error.ToDto());
        var details = error.Details;

        ShowError(title, message, details);
    }

    /// <summary>
    /// Shows a warning dialog.
    /// </summary>
    public static void ShowWarning(string title, string message)
    {
        MessageBox.Show(
            message,
            title,
            MessageBoxButton.OK,
            MessageBoxImage.Warning);
    }

    /// <summary>
    /// Shows an information dialog.
    /// </summary>
    public static void ShowInfo(string title, string message)
    {
        MessageBox.Show(
            message,
            title,
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    /// <summary>
    /// Shows a confirmation dialog and returns user's choice.
    /// </summary>
    public static bool ShowConfirmation(string title, string message)
    {
        var result = MessageBox.Show(
            message,
            title,
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        return result == MessageBoxResult.Yes;
    }

    /// <summary>
    /// Converts technical error messages to user-friendly messages.
    /// </summary>
    private static string GetUserFriendlyMessage(NormalizedErrorDto error)
    {
        // Map error codes to friendly messages
        return error.Code switch
        {
            // Authentication errors
            ErrorCode.AuthenticationFailed => "Authentication failed. Please check your credentials and try again.",

            ErrorCode.ConditionalAccessBlocked => "Access blocked by Conditional Access policy.\n\n" +
                                                 "Please contact your IT administrator for assistance.",

            ErrorCode.MfaRequired => "Multi-Factor Authentication is required.\n\n" +
                                   "Please complete the MFA verification in the authentication prompt.",

            ErrorCode.TokenExpired => "Your session has expired.\n\n" +
                                    "Please disconnect and reconnect to Exchange Online.",

            // Permission errors
            ErrorCode.PermissionDenied or ErrorCode.InsufficientPrivileges =>
                "You don't have permission to perform this operation.\n\n" +
                "Please contact your Exchange administrator to request the necessary permissions.",

            // Operation errors
            ErrorCode.CmdletNotAvailable => "The required PowerShell command is not available.\n\n" +
                                          "Please ensure the ExchangeOnlineManagement module is properly installed.",

            ErrorCode.ModuleNotLoaded => "The ExchangeOnlineManagement module is not loaded.\n\n" +
                                       "Try restarting the worker or reinstalling the module.",

            ErrorCode.InvalidParameter => $"Invalid parameter in the operation.\n\n{error.Message}",

            ErrorCode.OperationNotSupported => "This operation is not supported.\n\n" +
                                             "The feature may not be available with your Exchange configuration.",

            // Transient errors
            ErrorCode.Throttling => "The Exchange server is limiting the number of requests.\n\n" +
                                  "Please wait a moment. The operation will be retried automatically." +
                                  (error.RetryAfterSeconds.HasValue ? $"\n\nRetry after: {error.RetryAfterSeconds} seconds" : ""),

            ErrorCode.ServiceUnavailable => "The Exchange Online service is temporarily unavailable.\n\n" +
                                          "Please try again in a few moments.",

            ErrorCode.NetworkError => "A network error occurred.\n\n" +
                                    "Please check your internet connection and try again.",

            ErrorCode.Timeout => "The operation timed out.\n\n" +
                               "This may be due to network issues or server load. Please try again.",

            // Resource errors
            ErrorCode.ResourceNotFound => $"The requested resource was not found.\n\n{error.Message}",

            ErrorCode.ResourceAlreadyExists => $"The resource already exists.\n\n{error.Message}",

            // Worker errors
            ErrorCode.WorkerNotRunning => "The background worker is not running.\n\n" +
                                        "Please start the worker before connecting to Exchange Online.",

            ErrorCode.WorkerCrashed => "The background worker has stopped unexpectedly.\n\n" +
                                     "Try restarting the worker. Check the logs for more details.",

            ErrorCode.IpcError => "Communication error with the background worker.\n\n" +
                                "Try restarting the worker or the application.",

            // Unknown
            ErrorCode.Unknown or _ => !string.IsNullOrEmpty(error.Message)
                                    ? $"An error occurred:\n\n{error.Message}"
                                    : "An unexpected error occurred. Please check the logs for more information."
        };
    }
}
