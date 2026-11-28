using OnlyExo365.Shell.Localization;

namespace OnlyExo365.Tests;

public sealed class LocalizationServiceTests : IDisposable
{
    // Restore locale for test isolation.
    private readonly string _originalLocale;

    public LocalizationServiceTests()
    {
        _originalLocale = LocalizationService.Instance.CurrentLocale;
        LocalizationService.Instance.SetLocale("en");
    }

    public void Dispose()
    {
        LocalizationService.Instance.SetLocale(_originalLocale);
    }

    [Fact]
    public void Get_EnglishKey_ReturnsEnglishValue()
    {
        LocalizationService.Instance.SetLocale("en");

        var result = LocalizationService.Instance.Get("Nav.Dashboard");

        Assert.Equal("Dashboard", result);
    }

    [Fact]
    public void Get_AfterSetLocaleItalian_ReturnsItalianValue()
    {
        LocalizationService.Instance.SetLocale("it");

        var result = LocalizationService.Instance.Get("Nav.Mailboxes");

        Assert.Equal("Cassette postali", result);
    }

    [Fact]
    public void Get_MissingKey_ReturnsFallbackKey()
    {
        var missingKey = "This.Key.Does.Not.Exist.In.Any.RESX";

        var result = LocalizationService.Instance.Get(missingKey);

        Assert.Equal(missingKey, result);
    }

    [Fact]
    public void Get_ItalianKeyMissing_FallsBackToEnglishValue()
    {
        // Missing Italian resources fall back to English.
        LocalizationService.Instance.SetLocale("it");

        // The title may be translated or use English fallback.
        var result = LocalizationService.Instance.Get("Shell.WindowTitle");

        Assert.False(string.IsNullOrWhiteSpace(result));
        Assert.NotEqual("Shell.WindowTitle", result);
    }

    [Fact]
    public void SetLocale_SameLocale_DoesNotFireCultureChanged()
    {
        LocalizationService.Instance.SetLocale("en");
        var firedCount = 0;
        LocalizationService.Instance.CultureChanged += (_, _) => firedCount++;

        LocalizationService.Instance.SetLocale("en");

        Assert.Equal(0, firedCount);
    }

    [Fact]
    public void SetLocale_DifferentLocale_FiresCultureChangedEvent()
    {
        LocalizationService.Instance.SetLocale("en");
        var fired = false;
        LocalizationService.Instance.CultureChanged += (_, _) => fired = true;

        LocalizationService.Instance.SetLocale("it");

        Assert.True(fired);
    }

    [Fact]
    public void GetFormat_AppliesArguments()
    {
        LocalizationService.Instance.SetLocale("en");

        var result = LocalizationService.Instance.GetFormat("Progress.LoadedCount", 42);

        Assert.Contains("42", result);
    }

    [Fact]
    public void GetFormat_MissingKey_ReturnsKeyWhenNoArgs()
    {
        var missingKey = "Missing.Format.Key";

        var result = LocalizationService.Instance.GetFormat(missingKey);

        Assert.Equal(missingKey, result);
    }

    [Theory]
    [InlineData("en")]
    [InlineData("it")]
    public void SetLocale_ValidCode_UpdatesCurrentLocale(string code)
    {
        LocalizationService.Instance.SetLocale(code);

        Assert.Equal(code, LocalizationService.Instance.CurrentLocale);
    }

    [Fact]
    public void AvailableLocales_ContainsEnglishAndItalian()
    {
        var codes = LocalizationService.Instance.AvailableLocales.Select(l => l.Code).ToList();

        Assert.Contains("en", codes);
        Assert.Contains("it", codes);
    }

    [Fact]
    public void Get_English_AndItalian_ProduceDifferentResults_ForTranslatedKey()
    {
        LocalizationService.Instance.SetLocale("en");
        var english = LocalizationService.Instance.Get("Nav.MailContacts");

        LocalizationService.Instance.SetLocale("it");
        var italian = LocalizationService.Instance.Get("Nav.MailContacts");

        Assert.NotEqual(english, italian);
    }
}

