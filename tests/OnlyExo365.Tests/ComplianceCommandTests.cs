using System.Management.Automation;
using OnlyExo365.Contracts;
using OnlyExo365.Worker.PowerShell;

namespace OnlyExo365.Tests;

public sealed class ComplianceCommandTests
{
    [Fact]
    public void BuildGetComplianceWorkspaceScript_GuardsNullValuesAndEmitsStructuredWarnings()
    {
        var script = ExoComplianceCommands.BuildGetComplianceWorkspaceScript(25);

        Assert.Contains("if ($null -eq $item) { continue }", script, StringComparison.Ordinal);
        Assert.Contains("Write-Warning \"__EA_WARN__", script, StringComparison.Ordinal);
        Assert.Contains("Get-CaseHoldPolicy is not available", script, StringComparison.Ordinal);
        Assert.Contains("Get-FallbackHoldEntries", script, StringComparison.Ordinal);
        Assert.Contains("InPlaceHolds", script, StringComparison.Ordinal);
        Assert.Contains("[PSCustomObject]@{", script, StringComparison.Ordinal);
    }

    [Fact]
    public void MapWorkspaceResponse_ParsesPartialWorkspaceAndStructuredWarnings()
    {
        var workspace = CreatePsCustomObject(
            ("Searches", new object[]
            {
                CreatePsCustomObject(
                    ("Name", "Search-01"),
                    ("CaseName", "Case-01"),
                    ("Status", "Completed"),
                    ("ExchangeLocations", new[] { "user@contoso.com", null }),
                    ("ContentMatchQuery", "kind:email"))
            }),
            ("Cases", new object[]
            {
                CreatePsCustomObject(
                    ("Name", "Case-01"),
                    ("Status", "Active"),
                    ("CaseType", "eDiscovery"))
            }),
            ("Actions", new object[]
            {
                CreatePsCustomObject(
                    ("Identity", "hold-01"),
                    ("Name", "Hold-01"),
                    ("ActionType", "Hold"),
                    ("CaseName", "Case-01"),
                    ("ExchangeLocations", new[] { "user@contoso.com" }))
            }));

        var response = ExoComplianceCommands.MapWorkspaceResponse(
            [workspace],
            ["__EA_WARN__{\"code\":\"CaseHoldPolicyUnavailable\",\"scope\":\"ComplianceWorkspace.Holds\",\"message\":\"Existing holds are not visible in this Purview session: Get-CaseHoldPolicy is not available. The workspace shows searches and cases, but hold listing remains unsupported.\",\"isPartialData\":true}"]);

        Assert.Single(response.Searches);
        Assert.Single(response.Cases);
        Assert.Single(response.Actions);
        Assert.True(response.HasPartialData);
        Assert.Single(response.Warnings);
        Assert.True(response.IsHoldListingUnsupported);
        Assert.Contains("hold", response.HoldListingStatusMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("Search-01", response.Searches[0].Name);
        Assert.Equal(["user@contoso.com"], response.Searches[0].ExchangeLocations);
        Assert.Equal("Hold", response.Actions[0].ActionType);
    }

    [Fact]
    public void MapWorkspaceResponse_PreservesFallbackHoldEntriesWithoutUnsupportedState()
    {
        var workspace = CreatePsCustomObject(
            ("Actions", new object[]
            {
                CreatePsCustomObject(
                    ("Identity", "UniH123"),
                    ("Name", "UniH123"),
                    ("ActionType", "Hold"),
                    ("Status", "FallbackDetected"),
                    ("ExchangeLocations", new[] { "user@contoso.com" }),
                    ("Details", "Fallback from mailbox InPlaceHolds"))
            }));

        var response = ExoComplianceCommands.MapWorkspaceResponse(
            [workspace],
            ["__EA_WARN__{\"code\":\"CaseHoldFallbackUsed\",\"scope\":\"ComplianceWorkspace.Holds\",\"message\":\"Get-CaseHoldPolicy is not available. Visible holds were reconstructed from mailbox InPlaceHolds values; case/search metadata may be missing.\",\"isPartialData\":true}"]);

        Assert.Single(response.Actions);
        Assert.True(response.HasPartialData);
        Assert.False(response.IsHoldListingUnsupported);
        Assert.Null(response.HoldListingStatusMessage);
        Assert.Equal("UniH123", response.Actions[0].Name);
        Assert.Equal(["user@contoso.com"], response.Actions[0].ExchangeLocations);
    }

    [Fact]
    public void BuildConnectComplianceSearchOnlyCommand_AddsEnableSearchOnlySessionWithoutChangingBaseCmdlet()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            AuthenticationMode = ExchangeAuthenticationMode.Interactive,
            ExchangeOrganization = "contoso.onmicrosoft.com"
        };

        var command = ExoCommands.BuildConnectComplianceSearchOnlyCommand(configuration);

        Assert.Contains("Connect-IPPSSession", command, StringComparison.Ordinal);
        Assert.Contains("-EnableSearchOnlySession", command, StringComparison.Ordinal);
    }

    [Fact]
    public void MapSearchUnifiedAuditLogResponse_PreservesRowsAndLimitWarning()
    {
        var response = ExoComplianceCommands.MapSearchUnifiedAuditLogResponse(
            [
                CreatePsCustomObject(
                    ("Identity", "audit-01"),
                    ("CreationDate", new DateTime(2026, 3, 12, 10, 0, 0, DateTimeKind.Utc)),
                    ("Operations", "MailItemsAccessed"),
                    ("ObjectId", "user@contoso.com"))
            ],
            maxResults: 1);

        Assert.Single(response.Results);
        Assert.Equal(1, response.TotalCount);
        Assert.Contains("Results limited", response.Warning, StringComparison.Ordinal);
        Assert.Equal("audit-01", response.Results[0].Identity);
    }

    private static PSObject CreatePsCustomObject(params (string Name, object? Value)[] properties)
    {
        var obj = new PSObject();
        foreach (var (name, value) in properties)
        {
            obj.Properties.Add(new PSNoteProperty(name, value));
        }

        return obj;
    }
}

