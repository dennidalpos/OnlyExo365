using ExchangeAdmin.Worker.PowerShell;

namespace ExchangeAdmin.Tests;

public partial class WorkerCommandTests
{
    [Theory]
    [InlineData(null, "Name")]
    [InlineData("", "Name")]
    [InlineData("Identity", "Identity")]
    [InlineData("Path", "Identity")]
    [InlineData("PrimarySmtpAddress", "PrimarySmtpAddress")]
    [InlineData("BadProperty", "Name")]
    public void NormalizePublicFolderSortProperty_PrefersReadableAlphabeticDefault(string? input, string expected)
    {
        Assert.Equal(expected, ExoPublicFolderCommands.NormalizePublicFolderSortProperty(input));
    }

    [Fact]
    public void BuildGetRoleGroupDetailsScript_UsesReadableMemberDisplayFallbacks()
    {
        var script = ExoPermissionCommands.BuildGetRoleGroupDetailsScript(new GetRoleGroupDetailsRequest
        {
            Identity = "Organization Management"
        });

        Assert.Contains("function Test-IsRoleGroupMemberTechnicalValue", script, StringComparison.Ordinal);
        Assert.Contains("function Resolve-RoleGroup", script, StringComparison.Ordinal);
        Assert.Contains("foreach ($propertyName in $propertyNames)", script, StringComparison.Ordinal);
        Assert.Contains("Test-RoleGroupMatchValue $property.Value $normalizedIdentity", script, StringComparison.Ordinal);
        Assert.Contains("@('Identity')", script, StringComparison.Ordinal);
        Assert.Contains("@('Guid')", script, StringComparison.Ordinal);
        Assert.Contains("@('DistinguishedName')", script, StringComparison.Ordinal);
        Assert.Contains("@('Name', 'DisplayName')", script, StringComparison.Ordinal);
        Assert.Contains("Multiple role groups matched identity '$normalizedIdentity' via $matchedProperties", script, StringComparison.Ordinal);
        Assert.Contains("function Get-ResolvedRoleGroupMembers", script, StringComparison.Ordinal);
        Assert.Contains("Get-RoleGroupMember -Identity $candidate -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("@('DisplayName', 'PrimarySmtpAddress', 'WindowsLiveID', 'UserPrincipalName', 'Alias', 'Name')", script, StringComparison.Ordinal);
        Assert.Contains("Get-Recipient -Identity $identity -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("Get-User -Identity $identity -ErrorAction Stop", script, StringComparison.Ordinal);
        Assert.Contains("DisplayName = Resolve-RoleGroupMemberDisplayName $_", script, StringComparison.Ordinal);
        Assert.Contains("Identity = if ($_.Identity) { $_.Identity.ToString() } else { Resolve-RoleGroupMemberDisplayName $_ }", script, StringComparison.Ordinal);
        Assert.Contains("Members = @($members | ForEach-Object {", script, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildGetRoleGroupDetailsScript_PrefersHostedIdentityBeforeFriendlyNameFallback()
    {
        var script = ExoPermissionCommands.BuildGetRoleGroupDetailsScript(new GetRoleGroupDetailsRequest
        {
            Identity = "FFO.extest.microsoft.com/Microsoft Exchange Hosted Organizations/BioferSpA.onmicrosoft.com/Configuration/CommunicationComplianceAnalysts"
        });

        var identityIndex = script.IndexOf("@('Identity')", StringComparison.Ordinal);
        var guidIndex = script.IndexOf("@('Guid')", StringComparison.Ordinal);
        var distinguishedNameIndex = script.IndexOf("@('DistinguishedName')", StringComparison.Ordinal);
        var friendlyNameIndex = script.IndexOf("@('Name', 'DisplayName')", StringComparison.Ordinal);

        Assert.NotEqual(-1, identityIndex);
        Assert.NotEqual(-1, guidIndex);
        Assert.NotEqual(-1, distinguishedNameIndex);
        Assert.NotEqual(-1, friendlyNameIndex);
        Assert.True(identityIndex < guidIndex);
        Assert.True(guidIndex < distinguishedNameIndex);
        Assert.True(distinguishedNameIndex < friendlyNameIndex);
        Assert.Contains("return Get-RoleGroup -Identity $normalizedIdentity -ErrorAction Stop", script, StringComparison.Ordinal);
    }
}
