using OnlyExo365.Worker.PowerShell;

namespace OnlyExo365.Tests;

public partial class WorkerCommandTests
{
    [Theory]
    [InlineData(0, 50, 51)]
    [InlineData(10, 25, 36)]
    public void BuildGetDistributionListsScript_UsesWindowedPreloadWithoutUnlimitedWhenSearchIsEmpty(int skip, int pageSize, int expectedWindowSize)
    {
        var script = ExoGroupCommands.BuildGetDistributionListsScript(
            skip,
            pageSize,
            escapedSearch: string.Empty,
            escapedFilterType: string.Empty,
            sortProperty: "DisplayName",
            sortDirection: string.Empty,
            includeDynamic: true,
            includeUnified: true,
            useWindowedLoad: true);

        Assert.Contains($"$pageWindowSize = {expectedWindowSize}", script, StringComparison.Ordinal);
        Assert.Contains("Get-DistributionGroup -ResultSize $pageWindowSize", script, StringComparison.Ordinal);
        Assert.Contains("Get-DynamicDistributionGroup -ResultSize $pageWindowSize", script, StringComparison.Ordinal);
        Assert.Contains("Get-UnifiedGroup -ResultSize $pageWindowSize", script, StringComparison.Ordinal);
        Assert.DoesNotContain("Get-DistributionGroup -ResultSize Unlimited", script, StringComparison.Ordinal);
        Assert.Contains("IsTotalCountExact = $false", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildGetDynamicGroupMembersScript_PrefersDedicatedCmdletWhenAvailable()
    {
        var script = ExoGroupCommands.BuildGetDynamicGroupMembersScript(
            "dynamic@contoso.com",
            skip: 25,
            pageSize: 50,
            useDedicatedCmdlet: true);

        Assert.Contains("Get-DynamicDistributionGroupMember -Identity 'dynamic@contoso.com' -ResultSize Unlimited", script, StringComparison.Ordinal);
        Assert.DoesNotContain("Get-Recipient @recipientPreviewParams", script, StringComparison.Ordinal);
        Assert.Contains("$pagedMembers = $allMembers | Select-Object -Skip 25 -First 50", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildGetDynamicGroupMembersScript_UsesRecipientPreviewFallbackWithRecipientContainer()
    {
        var script = ExoGroupCommands.BuildGetDynamicGroupMembersScript(
            "dynamic@contoso.com",
            skip: 0,
            pageSize: 250,
            useDedicatedCmdlet: false);

        Assert.Contains("$ddg = Get-DynamicDistributionGroup -Identity 'dynamic@contoso.com'", script, StringComparison.Ordinal);
        Assert.Contains("RecipientPreviewFilter = $ddg.RecipientFilter", script, StringComparison.Ordinal);
        Assert.Contains("$recipientPreviewParams['OrganizationalUnit'] = $ddg.RecipientContainer", script, StringComparison.Ordinal);
        Assert.Contains("$allMembers = Get-Recipient @recipientPreviewParams", script, StringComparison.Ordinal);
    }
}

