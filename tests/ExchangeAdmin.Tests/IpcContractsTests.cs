using ExchangeAdmin.Contracts;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Infrastructure.Ipc;

namespace ExchangeAdmin.Tests;

public class IpcContractsTests
{
    [Fact]
    public void IpcSessionContext_BuildsStableScopedPipeNames()
    {
        var first = new IpcSessionContext
        {
            UserScope = "S-1-5-21-1000",
            SessionId = 42
        };

        var second = new IpcSessionContext
        {
            UserScope = "S-1-5-21-1000",
            SessionId = 42
        };

        Assert.Equal(first.RequestPipeName, second.RequestPipeName);
        Assert.Equal(first.EventPipeName, second.EventPipeName);
        Assert.Contains("_42", first.RequestPipeName);
        Assert.NotEqual(first.RequestPipeName, first.EventPipeName);
    }

    [Fact]
    public void IpcSessionContext_MatchesOnlySameUserAndSession()
    {
        var current = new IpcSessionContext
        {
            UserScope = "user-a",
            SessionId = 7
        };

        Assert.True(current.Matches(new IpcSessionContext { UserScope = "user-a", SessionId = 7 }));
        Assert.False(current.Matches(new IpcSessionContext { UserScope = "user-b", SessionId = 7 }));
        Assert.False(current.Matches(new IpcSessionContext { UserScope = "user-a", SessionId = 8 }));
        Assert.False(current.Matches(null));
    }

    [Fact]
    public void JsonMessageSerializer_RoundTripsHandshakeRequest()
    {
        var request = new HandshakeRequest
        {
            SessionToken = "token-123",
            SessionId = 12,
            UserScope = "scope-a"
        };

        var json = JsonMessageSerializer.Serialize(request);
        var deserialized = JsonMessageSerializer.DeserializeMessage(json);

        var typed = Assert.IsType<HandshakeRequest>(deserialized);
        Assert.Equal(MessageType.HandshakeRequest, typed.Type);
        Assert.Equal("token-123", typed.SessionToken);
        Assert.Equal(12, typed.SessionId);
        Assert.Equal("scope-a", typed.UserScope);
    }

    [Fact]
    public void JsonMessageSerializer_RoundTripsResourceMailboxRequest()
    {
        var request = new UpsertResourceMailboxRequest
        {
            Identity = "room1@contoso.com",
            ResourceType = "Room",
            DisplayName = "Room 1",
            Alias = "room1",
            PrimarySmtpAddress = "room1@contoso.com",
            HiddenFromAddressListsEnabled = true,
            BookingSettings = new ResourceBookingSettingsDto
            {
                AutomateProcessing = "AutoAccept",
                BookingWindowInDays = 120,
                MaximumDurationInMinutes = 90,
                ResourceDelegates = new List<string> { "delegate1@contoso.com" }
            }
        };

        var json = JsonMessageSerializer.Serialize(request);
        var typed = JsonMessageSerializer.Deserialize<UpsertResourceMailboxRequest>(json);
        Assert.NotNull(typed);

        Assert.Equal("Room", typed!.ResourceType);
        Assert.True(typed.HiddenFromAddressListsEnabled);
        Assert.Equal(120, typed.BookingSettings.BookingWindowInDays);
        Assert.Equal("delegate1@contoso.com", Assert.Single(typed.BookingSettings.ResourceDelegates));
    }

    [Fact]
    public void RequestEnvelope_PreservesResourceOperationType()
    {
        var envelope = new RequestEnvelope
        {
            Operation = OperationType.GetResourceMailboxDetails,
            Payload = JsonMessageSerializer.ToJsonElement(new GetResourceMailboxDetailsRequest
            {
                Identity = "room1@contoso.com"
            })
        };

        var json = JsonMessageSerializer.Serialize(envelope);
        var typed = JsonMessageSerializer.Deserialize<RequestEnvelope>(json);
        Assert.NotNull(typed);

        Assert.Equal(OperationType.GetResourceMailboxDetails, typed!.Operation);
        var payload = JsonMessageSerializer.ExtractPayload<GetResourceMailboxDetailsRequest>(typed.Payload);
        Assert.Equal("room1@contoso.com", payload?.Identity);
    }

    [Fact]
    public void JsonMessageSerializer_RoundTripsMobileDeviceRequests()
    {
        var request = new SetMobileDeviceAccessStateRequest
        {
            MailboxIdentity = "user@contoso.com",
            DeviceId = "Appl12345",
            AccessState = "Blocked"
        };

        var json = JsonMessageSerializer.Serialize(request);
        var typed = JsonMessageSerializer.Deserialize<SetMobileDeviceAccessStateRequest>(json);
        Assert.NotNull(typed);

        Assert.Equal("user@contoso.com", typed!.MailboxIdentity);
        Assert.Equal("Appl12345", typed.DeviceId);
        Assert.Equal("Blocked", typed.AccessState);
    }

    [Fact]
    public void RequestEnvelope_PreservesMobileDeviceOperationType()
    {
        var envelope = new RequestEnvelope
        {
            Operation = OperationType.GetMobileDevices,
            Payload = JsonMessageSerializer.ToJsonElement(new GetMobileDevicesRequest
            {
                SearchQuery = "iphone",
                AccessState = "Allowed"
            })
        };

        var json = JsonMessageSerializer.Serialize(envelope);
        var typed = JsonMessageSerializer.Deserialize<RequestEnvelope>(json);
        Assert.NotNull(typed);

        Assert.Equal(OperationType.GetMobileDevices, typed!.Operation);
        var payload = JsonMessageSerializer.ExtractPayload<GetMobileDevicesRequest>(typed.Payload);
        Assert.Equal("iphone", payload?.SearchQuery);
        Assert.Equal("Allowed", payload?.AccessState);
    }

    [Fact]
    public void JsonMessageSerializer_RoundTripsMigrationBatchRequests()
    {
        var request = new CompleteMigrationBatchRequest
        {
            Identity = "Batch-01"
        };

        var json = JsonMessageSerializer.Serialize(request);
        var typed = JsonMessageSerializer.Deserialize<CompleteMigrationBatchRequest>(json);
        Assert.NotNull(typed);

        Assert.Equal("Batch-01", typed!.Identity);
    }

    [Fact]
    public void JsonMessageSerializer_RoundTripsMailboxProvisioningRequest()
    {
        var request = new GetMailboxProvisioningCandidatesRequest
        {
            SearchQuery = "mario",
            OnlyWithoutLicense = true,
            OnlyWithoutMail = false,
            PageSize = 25,
            Skip = 50
        };

        var json = JsonMessageSerializer.Serialize(request);
        var typed = JsonMessageSerializer.Deserialize<GetMailboxProvisioningCandidatesRequest>(json);
        Assert.NotNull(typed);

        Assert.Equal("mario", typed!.SearchQuery);
        Assert.True(typed.OnlyWithoutLicense);
        Assert.False(typed.OnlyWithoutMail);
        Assert.Equal(25, typed.PageSize);
        Assert.Equal(50, typed.Skip);
    }

    [Fact]
    public void RequestEnvelope_PreservesMailboxProvisioningOperationType()
    {
        var envelope = new RequestEnvelope
        {
            Operation = OperationType.GetMailboxProvisioningCandidates,
            Payload = JsonMessageSerializer.ToJsonElement(new GetMailboxProvisioningCandidatesRequest
            {
                SearchQuery = "rossi",
                OnlyWithoutLicense = true,
                OnlyWithoutMail = true
            })
        };

        var json = JsonMessageSerializer.Serialize(envelope);
        var typed = JsonMessageSerializer.Deserialize<RequestEnvelope>(json);
        Assert.NotNull(typed);

        Assert.Equal(OperationType.GetMailboxProvisioningCandidates, typed!.Operation);
        var payload = JsonMessageSerializer.ExtractPayload<GetMailboxProvisioningCandidatesRequest>(typed.Payload);
        Assert.Equal("rossi", payload?.SearchQuery);
        Assert.True(payload?.OnlyWithoutLicense);
        Assert.True(payload?.OnlyWithoutMail);
    }

    [Fact]
    public void RequestEnvelope_PreservesMigrationBatchOperationType()
    {
        var envelope = new RequestEnvelope
        {
            Operation = OperationType.GetMigrationBatchDetails,
            Payload = JsonMessageSerializer.ToJsonElement(new GetMigrationBatchDetailsRequest
            {
                Identity = "Batch-02"
            })
        };

        var json = JsonMessageSerializer.Serialize(envelope);
        var typed = JsonMessageSerializer.Deserialize<RequestEnvelope>(json);
        Assert.NotNull(typed);

        Assert.Equal(OperationType.GetMigrationBatchDetails, typed!.Operation);
        var payload = JsonMessageSerializer.ExtractPayload<GetMigrationBatchDetailsRequest>(typed.Payload);
        Assert.Equal("Batch-02", payload?.Identity);
    }

    [Fact]
    public void JsonMessageSerializer_RoundTripsSharingPolicyRequest()
    {
        var request = new UpsertSharingPolicyRequest
        {
            Identity = "Default Sharing Policy",
            Name = "Partner Sharing",
            Domains = new List<string> { "contoso.com: CalendarSharingFreeBusyDetail" },
            Enabled = true,
            MakeDefault = false
        };

        var json = JsonMessageSerializer.Serialize(request);
        var typed = JsonMessageSerializer.Deserialize<UpsertSharingPolicyRequest>(json);
        Assert.NotNull(typed);

        Assert.Equal("Partner Sharing", typed!.Name);
        Assert.True(typed.Enabled);
        Assert.Equal("contoso.com: CalendarSharingFreeBusyDetail", Assert.Single(typed.Domains));
    }

    [Fact]
    public void RequestEnvelope_PreservesAddressListOperationType()
    {
        var envelope = new RequestEnvelope
        {
            Operation = OperationType.UpsertAddressList,
            Payload = JsonMessageSerializer.ToJsonElement(new UpsertAddressListRequest
            {
                Identity = "\\All Users",
                Name = "All Users",
                IncludedRecipients = new List<string> { "MailboxUsers" }
            })
        };

        var json = JsonMessageSerializer.Serialize(envelope);
        var typed = JsonMessageSerializer.Deserialize<RequestEnvelope>(json);
        Assert.NotNull(typed);

        Assert.Equal(OperationType.UpsertAddressList, typed!.Operation);
        var payload = JsonMessageSerializer.ExtractPayload<UpsertAddressListRequest>(typed.Payload);
        Assert.Equal("\\All Users", payload?.Identity);
        Assert.Equal("MailboxUsers", Assert.Single(payload?.IncludedRecipients ?? []));
    }

    [Fact]
    public void RequestEnvelope_PreservesExplicitTimeoutMs()
    {
        var envelope = new RequestEnvelope
        {
            Operation = OperationType.GetMailboxes,
            TimeoutMs = IpcConstants.RequestTimeoutMs
        };

        var json = JsonMessageSerializer.Serialize(envelope);
        var typed = JsonMessageSerializer.Deserialize<RequestEnvelope>(json);
        Assert.NotNull(typed);

        Assert.Equal(IpcConstants.RequestTimeoutMs, typed!.TimeoutMs);
    }

    [Fact]
    public void JsonMessageSerializer_RoundTripsInstallModuleRequestWithExplicitTarget()
    {
        var request = new InstallModuleRequest
        {
            ModuleName = "PowerShell7",
            InstallTarget = "PowerShell7",
            PackageId = "Microsoft.PowerShell"
        };

        var json = JsonMessageSerializer.Serialize(request);
        var typed = JsonMessageSerializer.Deserialize<InstallModuleRequest>(json);
        Assert.NotNull(typed);

        Assert.Equal("PowerShell7", typed!.ModuleName);
        Assert.Equal("PowerShell7", typed.InstallTarget);
        Assert.Equal("Microsoft.PowerShell", typed.PackageId);
    }

    [Fact]
    public void WorkerClientRuntime_CreateRequestEnvelope_UsesUniformTimeoutForReadOperations()
    {
        var request = WorkerClientRuntime.CreateRequestEnvelope(
            OperationType.GetMobileDevices,
            new GetMobileDevicesRequest
            {
                SearchQuery = "iphone",
                AccessState = "Allowed"
            });

        Assert.Equal(OperationType.GetMobileDevices, request.Operation);
        Assert.Equal(IpcConstants.RequestTimeoutMs, request.TimeoutMs);

        var payload = JsonMessageSerializer.ExtractPayload<GetMobileDevicesRequest>(request.Payload);
        Assert.Equal("iphone", payload?.SearchQuery);
        Assert.Equal("Allowed", payload?.AccessState);
    }

    [Fact]
    public void JsonMessageSerializer_RoundTripsWorkerConsoleVisibilityRequestAndResponse()
    {
        var request = new SetWorkerConsoleVisibilityRequest
        {
            IsVisible = true
        };

        var requestJson = JsonMessageSerializer.Serialize(request);
        var typedRequest = JsonMessageSerializer.Deserialize<SetWorkerConsoleVisibilityRequest>(requestJson);
        Assert.NotNull(typedRequest);
        Assert.True(typedRequest!.IsVisible);

        var response = new SetWorkerConsoleVisibilityResponse
        {
            IsVisible = true,
            Message = "Worker console shown."
        };

        var responseJson = JsonMessageSerializer.Serialize(response);
        var typedResponse = JsonMessageSerializer.Deserialize<SetWorkerConsoleVisibilityResponse>(responseJson);
        Assert.NotNull(typedResponse);
        Assert.True(typedResponse!.IsVisible);
        Assert.Equal("Worker console shown.", typedResponse.Message);
    }

    [Fact]
    public void RequestEnvelope_PreservesWorkerConsoleVisibilityOperationType()
    {
        var envelope = new RequestEnvelope
        {
            Operation = OperationType.SetWorkerConsoleVisibility,
            Payload = JsonMessageSerializer.ToJsonElement(new SetWorkerConsoleVisibilityRequest
            {
                IsVisible = false
            })
        };

        var json = JsonMessageSerializer.Serialize(envelope);
        var typed = JsonMessageSerializer.Deserialize<RequestEnvelope>(json);
        Assert.NotNull(typed);

        Assert.Equal(OperationType.SetWorkerConsoleVisibility, typed!.Operation);
        var payload = JsonMessageSerializer.ExtractPayload<SetWorkerConsoleVisibilityRequest>(typed.Payload);
        Assert.False(payload?.IsVisible);
    }
}
