using System.Text.Json;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Contracts.Security;
using ExchangeAdmin.Infrastructure.Ipc;

namespace ExchangeAdmin.Tests;

public sealed class IpcSecretHandlingTests
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    [Fact]
    public void CreateRequestEnvelope_RemovesPlainTextPasswordFromIpcPayload()
    {
        using var prepared = PreparedIpcPayload.Create(new CreateMailboxRequest
        {
            DisplayName = "Mario Rossi",
            Alias = "mrossi",
            PrimarySmtpAddress = "mrossi@contoso.com",
            MailboxType = "User",
            Password = "Sup3rSecret!"
        });

        var envelope = WorkerClientRuntime.CreateRequestEnvelope(OperationType.CreateMailbox, prepared.Payload);
        var payloadJson = envelope.Payload?.GetRawText() ?? string.Empty;

        Assert.DoesNotContain("Sup3rSecret!", payloadJson, StringComparison.Ordinal);
        Assert.Contains("passwordSecret", payloadJson, StringComparison.Ordinal);

        var payload = JsonSerializer.Deserialize<CreateMailboxRequest>(payloadJson, JsonOptions);

        Assert.NotNull(payload?.PasswordSecret);
        Assert.Null(payload?.Password);
        Assert.True(ProtectedSecretStore.Exists(payload!.PasswordSecret));
    }

    [Fact]
    public void ProtectedSecretStore_ConsumeReturnsSecretAndDeletesBackingFile()
    {
        var reference = ProtectedSecretStore.Create("Sup3rSecret!");

        Assert.NotNull(reference);
        Assert.True(ProtectedSecretStore.Exists(reference));

        var secret = ProtectedSecretStore.Consume(reference);

        Assert.Equal("Sup3rSecret!", secret);
        Assert.False(ProtectedSecretStore.Exists(reference));
    }
}
