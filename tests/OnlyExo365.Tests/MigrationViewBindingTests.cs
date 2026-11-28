namespace OnlyExo365.Tests;

public sealed class MigrationViewBindingTests
{
    [Theory]
    [InlineData("{loc:Loc Key=Migration.RefreshEndpoints}", "Migration.RefreshEndpointsCommand")]
    [InlineData("{loc:Loc Key=Migration.RunPreflight}", "Migration.RunBatchPreflightCommand")]
    [InlineData("{loc:Loc Key=Migration.CreateBatch}", "Migration.CreateBatchCommand")]
    public void MigrationView_ExposesEndpointAndBatchCommands(string buttonContent, string commandBinding)
    {
        var document = LoadViewDocument();

        var button = document
            .Descendants()
            .FirstOrDefault(candidate =>
                candidate.Name.LocalName == "Button" &&
                string.Equals((string?)candidate.Attribute("Content"), buttonContent, StringComparison.Ordinal) &&
                string.Equals((string?)candidate.Attribute("Command"), $"{{Binding {commandBinding}}}", StringComparison.Ordinal));

        Assert.NotNull(button);
    }

    [Fact]
    public void MigrationView_UsesPasswordBoxForEndpointCredentials()
    {
        var content = File.ReadAllText(GetViewPath());

        Assert.Contains("PasswordChanged=\"EndpointPasswordBox_OnPasswordChanged\"", content, StringComparison.Ordinal);
        Assert.Contains("helpers:PasswordBoxClearHelper.ClearTrigger=\"{Binding Migration.EndpointPasswordClearTrigger}\"", content, StringComparison.Ordinal);
        Assert.DoesNotContain("Text=\"{Binding Migration.EndpointPassword, UpdateSourceTrigger=PropertyChanged}\"", content, StringComparison.Ordinal);
    }

    private static System.Xml.Linq.XDocument LoadViewDocument()
    {
        return System.Xml.Linq.XDocument.Load(GetViewPath());
    }

    private static string GetViewPath()
    {
        return TestPathHelper.GetRepositoryPath("src", "OnlyExo365.Shell", "Views", "MigrationView.xaml");
    }
}

