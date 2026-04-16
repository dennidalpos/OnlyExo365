using ExchangeAdmin.Application.Security;
using ExchangeAdmin.Contracts;
using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Tests;

public sealed class LeastPrivilegeEvaluatorTests
{
    [Theory]
    [InlineData(ExchangeAuthenticationMode.Interactive)]
    [InlineData(ExchangeAuthenticationMode.DeviceCode)]
    [InlineData(ExchangeAuthenticationMode.AppCertificate)]
    [InlineData(ExchangeAuthenticationMode.ManagedIdentity)]
    public void EvaluateAll_MarksEveryLeastPrivilegeFeatureAvailableWhenAuthenticationAndCapabilitiesMatch(ExchangeAuthenticationMode authenticationMode)
    {
        var configuration = CreateConfiguration(authenticationMode);
        var evaluator = new LeastPrivilegeEvaluator(configuration);

        var evaluations = evaluator.EvaluateAll(CreateCapabilitiesForCatalog());

        Assert.Equal(LeastPrivilegeCatalog.All.Count, evaluations.Count);
        Assert.All(evaluations, evaluation => Assert.Equal(LeastPrivilegeFeatureStatus.Available, evaluation.Status));
    }

    [Fact]
    public void EvaluateAll_BlocksGraphFeatureWhenRequiredScopeIsMissing()
    {
        var configuration = new ExchangeOnlineConfiguration
        {
            GraphScopes = ["Organization.Read.All", "Directory.Read.All", "User.Read.All"]
        };

        var evaluator = new LeastPrivilegeEvaluator(configuration);

        var readEvaluation = evaluator.Evaluate(LeastPrivilegeCatalog.MailboxLicensingRead, new CapabilityMapDto());
        var writeEvaluation = evaluator.Evaluate(LeastPrivilegeCatalog.MailboxLicensingWrite, new CapabilityMapDto());

        Assert.Equal(LeastPrivilegeFeatureStatus.Blocked, readEvaluation.Status);
        Assert.DoesNotContain("LicenseAssignment.ReadWrite.All", readEvaluation.MissingRequirementsDisplay, StringComparison.Ordinal);
        Assert.Contains("RoleManagement.Read.Directory", readEvaluation.MissingRequirementsDisplay, StringComparison.Ordinal);

        Assert.Equal(LeastPrivilegeFeatureStatus.Blocked, writeEvaluation.Status);
        Assert.Contains("LicenseAssignment.ReadWrite.All", writeEvaluation.MissingRequirementsDisplay, StringComparison.Ordinal);
        Assert.Contains("RoleManagement.Read.Directory", writeEvaluation.MissingRequirementsDisplay, StringComparison.Ordinal);
    }

    [Fact]
    public void EvaluateAll_RequiresAdditionalSessionForComplianceWhenDedicatedCmdletsAreNotYetLoaded()
    {
        var evaluator = new LeastPrivilegeEvaluator(ExchangeOnlineConfiguration.CreateDefault());

        var evaluation = evaluator.Evaluate(
            LeastPrivilegeCatalog.ComplianceAuditAndEDiscovery,
            CreateCapabilities());

        Assert.Equal(LeastPrivilegeFeatureStatus.NeedsAdditionalSession, evaluation.Status);
        Assert.Contains("secondary session", evaluation.ValidationMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void EvaluateAll_AcceptsMailSecurityFeatureWhenAnySupportedSurfaceIsAvailable()
    {
        var evaluator = new LeastPrivilegeEvaluator(ExchangeOnlineConfiguration.CreateDefault());

        var evaluation = evaluator.Evaluate(
            LeastPrivilegeCatalog.MailSecurityBaseline,
            CreateCapabilities(("Get-DkimSigningConfig", true, null)));

        Assert.Equal(LeastPrivilegeFeatureStatus.Available, evaluation.Status);
    }

    [Fact]
    public void EvaluateAll_DefaultConfigurationKeepsGraphReadAvailableButLicenseWriteBlocked()
    {
        var evaluator = new LeastPrivilegeEvaluator(ExchangeOnlineConfiguration.CreateDefault());

        var readEvaluation = evaluator.Evaluate(
            LeastPrivilegeCatalog.MailboxLicensingRead,
            CreateCapabilitiesForCatalog());
        var writeEvaluation = evaluator.Evaluate(
            LeastPrivilegeCatalog.MailboxLicensingWrite,
            CreateCapabilitiesForCatalog());

        Assert.Equal(LeastPrivilegeFeatureStatus.Available, readEvaluation.Status);
        Assert.Equal(LeastPrivilegeFeatureStatus.Blocked, writeEvaluation.Status);
        Assert.Contains("LicenseAssignment.ReadWrite.All", writeEvaluation.MissingRequirementsDisplay, StringComparison.Ordinal);
    }

    [Fact]
    public void EvaluateAll_BlocksGraphFeatureForAppCertificateWithoutTenantId()
    {
        var evaluator = new LeastPrivilegeEvaluator(new ExchangeOnlineConfiguration
        {
            AuthenticationMode = ExchangeAuthenticationMode.AppCertificate,
            ApplicationId = "app-id",
            ExchangeOrganization = "tenant.onmicrosoft.com",
            CertificateThumbprint = "ABC123",
            GraphScopes =
            [
                "Organization.Read.All",
                "Directory.Read.All",
                "RoleManagement.Read.Directory",
                "User.Read.All"
            ],
            GraphLicenseWriteScopes =
            [
                "LicenseAssignment.ReadWrite.All"
            ]
        });

        var evaluation = evaluator.Evaluate(
            LeastPrivilegeCatalog.MailboxLicensingWrite,
            CreateCapabilitiesForCatalog());

        Assert.Equal(LeastPrivilegeFeatureStatus.Blocked, evaluation.Status);
        Assert.Contains("GraphTenantId", evaluation.MissingRequirementsDisplay, StringComparison.Ordinal);
    }

    [Fact]
    public void EvaluateAll_BlocksFeatureWhenAnyMandatoryCmdletIsUnavailable()
    {
        var evaluator = new LeastPrivilegeEvaluator(ExchangeOnlineConfiguration.CreateDefault());

        var evaluation = evaluator.Evaluate(
            LeastPrivilegeCatalog.MobileDevicesInventory,
            CreateCapabilitiesForCatalog(("Clear-MobileDevice", false, "RBAC denied")));

        Assert.Equal(LeastPrivilegeFeatureStatus.Blocked, evaluation.Status);
        Assert.Contains("Clear-MobileDevice", evaluation.MissingRequirementsDisplay, StringComparison.Ordinal);
        Assert.Contains("RBAC denied", evaluation.MissingRequirementsDisplay, StringComparison.Ordinal);
    }

    [Fact]
    public void EvaluateAll_AcceptsAlternativeCmdletFeatureWhenLegacyMessageTraceCmdletIsAvailable()
    {
        var evaluator = new LeastPrivilegeEvaluator(ExchangeOnlineConfiguration.CreateDefault());
        var capabilities = CreateCapabilitiesForCatalog();
        capabilities.Cmdlets["Get-MessageTraceV2"] = new CmdletCapabilityDto
        {
            Name = "Get-MessageTraceV2",
            IsAvailable = false,
            UnavailableReason = "Module version too old"
        };
        capabilities.Cmdlets["Get-MessageTrace"] = new CmdletCapabilityDto
        {
            Name = "Get-MessageTrace",
            IsAvailable = true
        };

        var evaluation = evaluator.Evaluate(
            LeastPrivilegeCatalog.MessageTraceRead,
            capabilities);

        Assert.Equal(LeastPrivilegeFeatureStatus.Available, evaluation.Status);
    }

    private static CapabilityMapDto CreateCapabilities(params (string Name, bool IsAvailable, string? Reason)[] cmdlets)
    {
        var map = new CapabilityMapDto();
        foreach (var (name, isAvailable, reason) in cmdlets)
        {
            map.Cmdlets[name] = new CmdletCapabilityDto
            {
                Name = name,
                IsAvailable = isAvailable,
                UnavailableReason = reason
            };
        }

        return map;
    }

    private static ExchangeOnlineConfiguration CreateConfiguration(ExchangeAuthenticationMode authenticationMode)
    {
        return new ExchangeOnlineConfiguration
        {
            AuthenticationMode = authenticationMode,
            ExchangeOrganization = "tenant.onmicrosoft.com",
            ApplicationId = "app-id",
            CertificateThumbprint = "ABC123",
            GraphTenantId = "graph-tenant-id",
            GraphScopes =
            [
                "Organization.Read.All",
                "Directory.Read.All",
                "RoleManagement.Read.Directory",
                "User.Read.All"
            ],
            GraphLicenseWriteScopes =
            [
                "LicenseAssignment.ReadWrite.All"
            ]
        };
    }

    private static CapabilityMapDto CreateCapabilitiesForCatalog(params (string Name, bool IsAvailable, string? Reason)[] overrides)
    {
        var capabilities = new CapabilityMapDto();
        var overrideMap = overrides.ToDictionary(item => item.Name, StringComparer.Ordinal);
        var cmdletNames = LeastPrivilegeCatalog.All
            .SelectMany(feature => feature.RequiredCmdletsAll.Concat(feature.RequiredCmdletsAny))
            .Distinct(StringComparer.Ordinal);

        foreach (var cmdletName in cmdletNames)
        {
            if (overrideMap.TryGetValue(cmdletName, out var cmdletOverride))
            {
                capabilities.Cmdlets[cmdletName] = new CmdletCapabilityDto
                {
                    Name = cmdletName,
                    IsAvailable = cmdletOverride.IsAvailable,
                    UnavailableReason = cmdletOverride.Reason
                };
            }
            else
            {
                capabilities.Cmdlets[cmdletName] = new CmdletCapabilityDto
                {
                    Name = cmdletName,
                    IsAvailable = true
                };
            }
        }

        return capabilities;
    }
}
