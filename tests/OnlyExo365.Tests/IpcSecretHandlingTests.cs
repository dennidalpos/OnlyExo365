using System.Text.Json;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Security;
using OnlyExo365.Shell.Ipc;

namespace OnlyExo365.Tests;

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

    [Theory]
    [InlineData("..\\outside")]
    [InlineData("secret")]
    [InlineData("11111111-1111-1111-1111-111111111111")]
    public void ProtectedSecretStore_RejectsMalformedSecretReferenceIds(string id)
    {
        var reference = new ProtectedSecretReference { Id = id };

        Assert.False(ProtectedSecretStore.Exists(reference));
        Assert.Null(ProtectedSecretStore.Consume(reference));

        ProtectedSecretStore.TryDelete(reference);
    }
}

