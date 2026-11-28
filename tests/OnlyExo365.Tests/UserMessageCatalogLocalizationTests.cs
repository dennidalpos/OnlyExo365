using OnlyExo365.Contracts.Messages;
using OnlyExo365.Shell.Localization;
using OnlyExo365.Shell.Text;

namespace OnlyExo365.Tests;

public sealed class UserMessageCatalogLocalizationTests : IDisposable
{
    private readonly string _originalLocale;

    public UserMessageCatalogLocalizationTests()
    {
        _originalLocale = LocalizationService.Instance.CurrentLocale;
        LocalizationService.Instance.SetLocale("en");
    }

    public void Dispose()
    {
        LocalizationService.Instance.SetLocale(_originalLocale);
    }

    [Theory]
    [InlineData(ErrorCode.AuthenticationFailed)]
    [InlineData(ErrorCode.ConditionalAccessBlocked)]
    [InlineData(ErrorCode.MfaRequired)]
    [InlineData(ErrorCode.TokenExpired)]
    [InlineData(ErrorCode.PermissionDenied)]
    [InlineData(ErrorCode.InsufficientPrivileges)]
    [InlineData(ErrorCode.CmdletNotAvailable)]
    [InlineData(ErrorCode.ModuleNotLoaded)]
    [InlineData(ErrorCode.InvalidParameter)]
    [InlineData(ErrorCode.OperationNotSupported)]
    [InlineData(ErrorCode.Throttling)]
    [InlineData(ErrorCode.ServiceUnavailable)]
    [InlineData(ErrorCode.NetworkError)]
    [InlineData(ErrorCode.Timeout)]
    [InlineData(ErrorCode.ResourceNotFound)]
    [InlineData(ErrorCode.ResourceAlreadyExists)]
    [InlineData(ErrorCode.WorkerNotRunning)]
    [InlineData(ErrorCode.WorkerCrashed)]
    [InlineData(ErrorCode.IpcError)]
    [InlineData(ErrorCode.Unknown)]
    public void GetFriendlyErrorMessage_AllErrorCodes_ReturnsNonEmptyString(ErrorCode code)
    {
        LocalizationService.Instance.SetLocale("en");
        var error = new NormalizedErrorDto { Code = code, Message = "test detail" };

        var result = UserMessageCatalog.GetFriendlyErrorMessage(error);

        Assert.False(string.IsNullOrWhiteSpace(result),
            $"Expected non-empty message for ErrorCode.{code}");
    }

    [Theory]
    [InlineData(ErrorCode.AuthenticationFailed)]
    [InlineData(ErrorCode.TokenExpired)]
    [InlineData(ErrorCode.WorkerNotRunning)]
    [InlineData(ErrorCode.NetworkError)]
    [InlineData(ErrorCode.PermissionDenied)]
    public void GetFriendlyErrorMessage_Italian_DiffersFromEnglish(ErrorCode code)
    {
        var error = new NormalizedErrorDto { Code = code, Message = "detail" };

        LocalizationService.Instance.SetLocale("en");
        var english = UserMessageCatalog.GetFriendlyErrorMessage(error);

        LocalizationService.Instance.SetLocale("it");
        var italian = UserMessageCatalog.GetFriendlyErrorMessage(error);

        Assert.True(english != italian,
            $"Expected Italian message for {code} to differ from English");
    }

    [Fact]
    public void GetFriendlyErrorMessage_Throttling_WithRetryAfter_IncludesSeconds()
    {
        LocalizationService.Instance.SetLocale("en");
        var error = new NormalizedErrorDto
        {
            Code = ErrorCode.Throttling,
            RetryAfterSeconds = 30
        };

        var result = UserMessageCatalog.GetFriendlyErrorMessage(error);

        Assert.Contains("30", result);
    }

    [Fact]
    public void GetFriendlyErrorMessage_Throttling_WithoutRetryAfter_DoesNotThrow()
    {
        var error = new NormalizedErrorDto { Code = ErrorCode.Throttling };

        var result = UserMessageCatalog.GetFriendlyErrorMessage(error);

        Assert.False(string.IsNullOrWhiteSpace(result));
    }

    [Fact]
    public void FormatMutationConfirmation_WithImpact_IncludesAllFields()
    {
        LocalizationService.Instance.SetLocale("en");

        var result = UserMessageCatalog.FormatMutationConfirmation("Delete", "user@contoso.com", "Permanent");

        Assert.Contains("Delete", result);
        Assert.Contains("user@contoso.com", result);
        Assert.Contains("Permanent", result);
    }

    [Fact]
    public void FormatMutationConfirmation_WithoutImpact_OmitsImpactText()
    {
        LocalizationService.Instance.SetLocale("en");

        var withImpact = UserMessageCatalog.FormatMutationConfirmation("Op", "target", "some impact");
        var withoutImpact = UserMessageCatalog.FormatMutationConfirmation("Op", "target");

        Assert.True(withImpact.Length > withoutImpact.Length,
            "Message with impact should be longer than message without impact");
        Assert.DoesNotContain("some impact", withoutImpact);
    }

    [Fact]
    public void StaticProperties_English_ReturnNonEmptyStrings()
    {
        LocalizationService.Instance.SetLocale("en");

        Assert.False(string.IsNullOrWhiteSpace(UserMessageCatalog.OperationInProgressTitle));
        Assert.False(string.IsNullOrWhiteSpace(UserMessageCatalog.OperationInProgressMessage));
        Assert.False(string.IsNullOrWhiteSpace(UserMessageCatalog.UnsavedChangesTitle));
        Assert.False(string.IsNullOrWhiteSpace(UserMessageCatalog.UnsavedChangesMessage));
        Assert.False(string.IsNullOrWhiteSpace(UserMessageCatalog.ConfirmOperationTitle));
        Assert.False(string.IsNullOrWhiteSpace(UserMessageCatalog.ConfirmOperationPrompt));
        Assert.False(string.IsNullOrWhiteSpace(UserMessageCatalog.ConnectionRequiredAlertTitle));
        Assert.False(string.IsNullOrWhiteSpace(UserMessageCatalog.ConnectionRequiredAlertMessage));
        Assert.False(string.IsNullOrWhiteSpace(UserMessageCatalog.LoadFailedAlertTitle));
        Assert.False(string.IsNullOrWhiteSpace(UserMessageCatalog.FormatPageUnavailableMessage("Dashboard")));
    }

    [Fact]
    public void StaticProperties_Italian_DifferFromEnglish()
    {
        LocalizationService.Instance.SetLocale("en");
        var englishTitle = UserMessageCatalog.UnsavedChangesTitle;

        LocalizationService.Instance.SetLocale("it");
        var italianTitle = UserMessageCatalog.UnsavedChangesTitle;
        var italianAlertTitle = UserMessageCatalog.ConnectionRequiredAlertTitle;

        Assert.NotEqual(englishTitle, italianTitle);
        Assert.NotEqual(UserMessageCatalog.LoadFailedAlertTitle, UserMessageCatalog.ConnectionRequiredAlertTitle);
        Assert.False(string.IsNullOrWhiteSpace(italianAlertTitle));
    }

    [Fact]
    public void CombineMessageAndDetails_WithDetails_IncludesBoth()
    {
        LocalizationService.Instance.SetLocale("en");

        var result = UserMessageCatalog.CombineMessageAndDetails("Main message", "Some details");

        Assert.Contains("Main message", result);
        Assert.Contains("Some details", result);
    }

    [Fact]
    public void CombineMessageAndDetails_WithNullDetails_ReturnsMessageOnly()
    {
        var result = UserMessageCatalog.CombineMessageAndDetails("Main message", null);

        Assert.Equal("Main message", result);
    }
}

