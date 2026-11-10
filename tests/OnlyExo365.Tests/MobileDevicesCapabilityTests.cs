using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Shell.ViewModels;
using OnlyExo365.Worker.PowerShell;

namespace OnlyExo365.Tests;

public sealed class MobileDevicesCapabilityTests
{
    [Fact]
    public void MobileDevicesCapabilityState_BlocksModuleWhenGetMobileDeviceIsUnavailable()
    {
        var capabilities = CreateCapabilities(features =>
        {
            features.CanGetMobileDevice = false;
        },
        ("Get-MobileDevice", false, "Cmdlet not found"));

        var state = MobileDevicesCapabilityState.From(capabilities);

        Assert.False(state.IsModuleAvailable);
        Assert.False(state.CanLoadPolicies);
        Assert.False(state.CanManageAccessState);
        Assert.False(state.CanRemoteWipe);
        Assert.NotNull(state.Message);
        Assert.Contains("Get-MobileDevice", state.Message, StringComparison.Ordinal);
        Assert.Contains("Cmdlet not found", state.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void MobileDevicesCapabilityState_ReportsReducedFunctionalityWhenPoliciesAndRemoteWipeAreUnavailable()
    {
        var capabilities = CreateCapabilities(features =>
        {
            features.CanGetMobileDevice = true;
            features.CanGetCasMailbox = true;
            features.CanSetCasMailbox = true;
            features.CanGetMobileDeviceMailboxPolicy = false;
            features.CanClearMobileDevice = false;
        },
        ("Get-MobileDeviceMailboxPolicy", false, "Role not assigned"),
        ("Clear-MobileDevice", false, "Tenant capability disabled"));

        var state = MobileDevicesCapabilityState.From(capabilities);

        Assert.True(state.IsModuleAvailable);
        Assert.False(state.CanLoadPolicies);
        Assert.True(state.CanManageAccessState);
        Assert.False(state.CanAssignPolicy);
        Assert.False(state.CanRemoteWipe);
        Assert.NotNull(state.Message);
        Assert.Contains("reduced functionality", state.Message, StringComparison.Ordinal);
        Assert.Contains("mailbox policy", state.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("remote wipe", state.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MobileDeviceCapabilityGuard_ReturnsExplicitReasonWhenListingIsUnavailable()
    {
        var capabilities = CreateCapabilities(features => features.CanGetMobileDevice = false,
            ("Get-MobileDevice", false, "The term 'Get-MobileDevice' is not recognized"));

        var exception = Assert.Throws<InvalidOperationException>(() => MobileDeviceCapabilityGuard.EnsureListingAvailable(capabilities));

        Assert.Contains("Get-MobileDevice", exception.Message, StringComparison.Ordinal);
        Assert.Contains("not recognized", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    private static CapabilityMapDto CreateCapabilities(
        Action<FeatureCapabilitiesDto> configureFeatures,
        params (string Name, bool IsAvailable, string? Reason)[] cmdlets)
    {
        var features = new FeatureCapabilitiesDto
        {
            CanGetMobileDevice = true
        };

        configureFeatures(features);

        var map = new CapabilityMapDto
        {
            Features = features
        };

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
}

