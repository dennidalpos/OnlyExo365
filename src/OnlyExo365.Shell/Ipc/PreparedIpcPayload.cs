using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Security;

namespace OnlyExo365.Shell.Ipc;

internal sealed class PreparedIpcPayload : IDisposable
{
    private readonly List<ProtectedSecretReference> _secretReferences = [];

    private PreparedIpcPayload(object? payload)
    {
        Payload = payload;
    }

    public object? Payload { get; }

    public static PreparedIpcPayload Create(object? payload)
    {
        if (payload == null)
        {
            return new PreparedIpcPayload(null);
        }

        return payload switch
        {
            UpsertContactRequest request => Create(request),
            CreateMailboxRequest request => Create(request),
            UpsertMigrationEndpointRequest request => Create(request),
            TestMigrationEndpointRequest request => Create(request),
            _ => new PreparedIpcPayload(payload)
        };
    }

    public void Dispose()
    {
        foreach (var reference in _secretReferences)
        {
            ProtectedSecretStore.TryDelete(reference);
        }
    }

    private static PreparedIpcPayload Create(UpsertContactRequest request)
    {
        var prepared = new PreparedIpcPayload(new UpsertContactRequest
        {
            Identity = request.Identity,
            ContactKind = request.ContactKind,
            DisplayName = request.DisplayName,
            Name = request.Name,
            Alias = request.Alias,
            PrimarySmtpAddress = request.PrimarySmtpAddress,
            ExternalEmailAddress = request.ExternalEmailAddress,
            UserPrincipalName = request.UserPrincipalName,
            PasswordSecret = ProtectedSecretStore.Create(request.Password),
            HiddenFromAddressListsEnabled = request.HiddenFromAddressListsEnabled
        });

        prepared.TrackSecret(((UpsertContactRequest)prepared.Payload!).PasswordSecret);
        return prepared;
    }

    private static PreparedIpcPayload Create(CreateMailboxRequest request)
    {
        var prepared = new PreparedIpcPayload(new CreateMailboxRequest
        {
            DisplayName = request.DisplayName,
            Alias = request.Alias,
            PrimarySmtpAddress = request.PrimarySmtpAddress,
            MailboxType = request.MailboxType,
            PasswordSecret = ProtectedSecretStore.Create(request.Password)
        });

        prepared.TrackSecret(((CreateMailboxRequest)prepared.Payload!).PasswordSecret);
        return prepared;
    }

    private static PreparedIpcPayload Create(UpsertMigrationEndpointRequest request)
    {
        var prepared = new PreparedIpcPayload(new UpsertMigrationEndpointRequest
        {
            Identity = request.Identity,
            Name = request.Name,
            EndpointType = request.EndpointType,
            RemoteServer = request.RemoteServer,
            RpcProxyServer = request.RpcProxyServer,
            ExchangeServer = request.ExchangeServer,
            EmailAddress = request.EmailAddress,
            RemoteTenant = request.RemoteTenant,
            Port = request.Port,
            Security = request.Security,
            Authentication = request.Authentication,
            Username = request.Username,
            PasswordSecret = ProtectedSecretStore.Create(request.Password),
            MaxConcurrentMigrations = request.MaxConcurrentMigrations,
            MaxConcurrentIncrementalSyncs = request.MaxConcurrentIncrementalSyncs,
            SkipVerification = request.SkipVerification,
            AcceptUntrustedCertificates = request.AcceptUntrustedCertificates
        });

        prepared.TrackSecret(((UpsertMigrationEndpointRequest)prepared.Payload!).PasswordSecret);
        return prepared;
    }

    private static PreparedIpcPayload Create(TestMigrationEndpointRequest request)
    {
        var prepared = new PreparedIpcPayload(new TestMigrationEndpointRequest
        {
            Identity = request.Identity,
            UseExistingEndpoint = request.UseExistingEndpoint,
            EndpointType = request.EndpointType,
            RemoteServer = request.RemoteServer,
            RpcProxyServer = request.RpcProxyServer,
            ExchangeServer = request.ExchangeServer,
            EmailAddress = request.EmailAddress,
            Port = request.Port,
            Security = request.Security,
            Authentication = request.Authentication,
            Username = request.Username,
            PasswordSecret = ProtectedSecretStore.Create(request.Password),
            SkipVerification = request.SkipVerification,
            AcceptUntrustedCertificates = request.AcceptUntrustedCertificates
        });

        prepared.TrackSecret(((TestMigrationEndpointRequest)prepared.Payload!).PasswordSecret);
        return prepared;
    }

    private void TrackSecret(ProtectedSecretReference? reference)
    {
        if (reference != null)
        {
            _secretReferences.Add(reference);
        }
    }
}

