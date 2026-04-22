using OnlyExo365.Contracts;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Worker.Ipc;

namespace OnlyExo365.Tests;

public class IpcServerOversizeTests
{
    [Fact]
    public void SerializeResponseForTransport_ReturnsFallbackErrorWhenResponseExceedsLimit()
    {
        var oversizedPayload = new string('x', IpcConstants.MaxMessageSizeBytes + 1024);
        var response = new ResponseEnvelope
        {
            CorrelationId = "corr-oversize",
            Success = true,
            Payload = JsonMessageSerializer.ToJsonElement(new { data = oversizedPayload })
        };

        var serialized = IpcServer.SerializeResponseForTransport(response);
        var fallback = JsonMessageSerializer.Deserialize<ResponseEnvelope>(serialized);

        Assert.NotNull(fallback);
        Assert.False(fallback!.Success);
        Assert.Equal("corr-oversize", fallback.CorrelationId);
        Assert.Equal(ErrorCode.MessageTooLarge, fallback.Error?.Code);
        Assert.Contains("maximum IPC message size", fallback.Error?.Message);
        Assert.True(IpcConstants.IsValidMessageSize(serialized.Length));
    }

    [Fact]
    public void CreateOversizeResponse_UsesExplicitIpcErrorContract()
    {
        var response = IpcServer.CreateOversizeResponse("corr-42", IpcConstants.MaxMessageSizeBytes + 1);

        Assert.False(response.Success);
        Assert.Equal("corr-42", response.CorrelationId);
        Assert.Equal(ErrorCode.MessageTooLarge, response.Error?.Code);
        Assert.Contains((IpcConstants.MaxMessageSizeBytes + 1).ToString(), response.Error?.Message);
    }
}

