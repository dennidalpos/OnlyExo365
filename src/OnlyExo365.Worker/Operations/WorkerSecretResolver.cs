using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Security;

namespace OnlyExo365.Worker.Operations;

internal static class WorkerSecretResolver
{
    public static void Resolve(UpsertContactRequest request)
    {
        request.Password = ProtectedSecretStore.Consume(request.PasswordSecret);
        request.PasswordSecret = null;
    }

    public static void Resolve(CreateMailboxRequest request)
    {
        request.Password = ProtectedSecretStore.Consume(request.PasswordSecret);
        request.PasswordSecret = null;
    }

    public static void Resolve(UpsertMigrationEndpointRequest request)
    {
        request.Password = ProtectedSecretStore.Consume(request.PasswordSecret);
        request.PasswordSecret = null;
    }

    public static void Resolve(TestMigrationEndpointRequest request)
    {
        request.Password = ProtectedSecretStore.Consume(request.PasswordSecret);
        request.PasswordSecret = null;
    }
}

