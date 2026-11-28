using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Errors;
using OnlyExo365.Shell.Localization;

namespace OnlyExo365.Shell.Text;

public static class UserMessageCatalog
{
    public static string DetailsHeading => Loc.Get("Msg.DetailsHeading");
    public static string OperationInProgressTitle => Loc.Get("Msg.OperationInProgressTitle");
    public static string OperationInProgressMessage => Loc.Get("Msg.OperationInProgressText");
    public static string UnsavedChangesTitle => Loc.Get("Msg.UnsavedChangesTitle");
    public static string UnsavedChangesMessage => Loc.Get("Msg.UnsavedChangesText");
    public static string ConfirmOperationTitle => Loc.Get("Msg.ConfirmOperationTitle");
    public static string ConfirmOperationPrompt => Loc.Get("Msg.ConfirmOperationPrompt");
    public static string ConnectionRequiredAlertTitle => Loc.Get("Alert.ConnectionRequiredTitle");
    public static string ConnectionRequiredAlertMessage => Loc.Get("Alert.ConnectionRequiredMessage");
    public static string LoadFailedAlertTitle => Loc.Get("Alert.LoadFailedTitle");

    public static string CombineMessageAndDetails(string message, string? details)
    {
        if (string.IsNullOrWhiteSpace(details))
        {
            return message;
        }

        return $"{message}{Environment.NewLine}{Environment.NewLine}{DetailsHeading}{Environment.NewLine}{details}";
    }

    public static string FormatMutationConfirmation(string operation, string target, string? impact = null)
    {
        return string.IsNullOrWhiteSpace(impact)
            ? Loc.GetFormat("Msg.ConfirmOperationFormat", operation, target)
            : Loc.GetFormat("Msg.ConfirmOperationImpactFormat", operation, target, impact);
    }

    public static string FormatPageUnavailableMessage(string pageTitle)
        => Loc.GetFormat("Alert.PageUnavailableMessageFormat", pageTitle);

    public static string GetFriendlyErrorMessage(NormalizedErrorDto error)
    {
        return error.Code switch
        {
            ErrorCode.AuthenticationFailed => Loc.Get("Error.AuthenticationFailed"),
            ErrorCode.ConditionalAccessBlocked => Loc.Get("Error.ConditionalAccessBlocked"),
            ErrorCode.MfaRequired => Loc.Get("Error.MfaRequired"),
            ErrorCode.TokenExpired => Loc.Get("Error.TokenExpired"),
            ErrorCode.PermissionDenied or ErrorCode.InsufficientPrivileges => Loc.Get("Error.PermissionDenied"),
            ErrorCode.CmdletNotAvailable => Loc.Get("Error.CmdletNotAvailable"),
            ErrorCode.ModuleNotLoaded => Loc.Get("Error.ModuleNotLoaded"),
            ErrorCode.InvalidParameter => Loc.GetFormat("Error.InvalidParameter", error.Message ?? string.Empty),
            ErrorCode.OperationNotSupported => Loc.Get("Error.OperationNotSupported"),
            ErrorCode.Throttling => error.RetryAfterSeconds.HasValue
                ? Loc.GetFormat("Error.ThrottlingRetryAfter", error.RetryAfterSeconds)
                : Loc.Get("Error.Throttling"),
            ErrorCode.ServiceUnavailable => Loc.Get("Error.ServiceUnavailable"),
            ErrorCode.NetworkError => Loc.Get("Error.NetworkError"),
            ErrorCode.Timeout => Loc.Get("Error.Timeout"),
            ErrorCode.ResourceNotFound => Loc.GetFormat("Error.ResourceNotFound", error.Message ?? string.Empty),
            ErrorCode.ResourceAlreadyExists => Loc.GetFormat("Error.ResourceAlreadyExists", error.Message ?? string.Empty),
            ErrorCode.WorkerNotRunning => Loc.Get("Error.WorkerNotRunning"),
            ErrorCode.WorkerCrashed => Loc.Get("Error.WorkerCrashed"),
            ErrorCode.IpcError => Loc.Get("Error.IpcError"),
            ErrorCode.Unknown or _ => !string.IsNullOrEmpty(error.Message)
                ? Loc.GetFormat("Error.UnknownWithMessage", error.Message)
                : Loc.Get("Error.UnknownGeneric")
        };
    }
}

