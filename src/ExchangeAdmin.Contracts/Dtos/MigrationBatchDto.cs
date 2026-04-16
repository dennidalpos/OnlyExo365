using System.Text.Json.Serialization;
using ExchangeAdmin.Contracts.Paging;
using ExchangeAdmin.Contracts.Security;

namespace ExchangeAdmin.Contracts.Dtos;

public class MigrationBatchListItemDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("status")]
    public string Status { get; set; } = string.Empty;

    [JsonPropertyName("state")]
    public string? State { get; set; }

    [JsonPropertyName("batchType")]
    public string? BatchType { get; set; }

    [JsonPropertyName("sourceEndpoint")]
    public string? SourceEndpoint { get; set; }

    [JsonPropertyName("targetEndpoint")]
    public string? TargetEndpoint { get; set; }

    [JsonPropertyName("totalCount")]
    public int? TotalCount { get; set; }

    [JsonPropertyName("activeCount")]
    public int? ActiveCount { get; set; }

    [JsonPropertyName("syncedCount")]
    public int? SyncedCount { get; set; }

    [JsonPropertyName("finalizedCount")]
    public int? FinalizedCount { get; set; }

    [JsonPropertyName("failedCount")]
    public int? FailedCount { get; set; }

    [JsonPropertyName("stoppedCount")]
    public int? StoppedCount { get; set; }

    [JsonPropertyName("createdBy")]
    public string? CreatedBy { get; set; }

    [JsonPropertyName("createdDateTime")]
    public DateTime? CreatedDateTime { get; set; }

    [JsonPropertyName("startDateTime")]
    public DateTime? StartDateTime { get; set; }

    [JsonPropertyName("completeDateTime")]
    public DateTime? CompleteDateTime { get; set; }

    [JsonPropertyName("lastSyncedDateTime")]
    public DateTime? LastSyncedDateTime { get; set; }
}

public class MigrationBatchDetailsDto : MigrationBatchListItemDto
{
    [JsonPropertyName("statusDetail")]
    public string? StatusDetail { get; set; }

    [JsonPropertyName("notificationEmails")]
    public List<string> NotificationEmails { get; set; } = new();

    [JsonPropertyName("autoStart")]
    public bool? AutoStart { get; set; }

    [JsonPropertyName("autoComplete")]
    public bool? AutoComplete { get; set; }

    [JsonPropertyName("badItemLimit")]
    public int? BadItemLimit { get; set; }

    [JsonPropertyName("largeItemLimit")]
    public int? LargeItemLimit { get; set; }

    [JsonPropertyName("report")]
    public string? Report { get; set; }

    [JsonPropertyName("startAfter")]
    public DateTime? StartAfter { get; set; }

    [JsonPropertyName("completeAfter")]
    public DateTime? CompleteAfter { get; set; }
}

public class MigrationEndpointDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("endpointType")]
    public string EndpointType { get; set; } = string.Empty;

    [JsonPropertyName("remoteServer")]
    public string? RemoteServer { get; set; }

    [JsonPropertyName("rpcProxyServer")]
    public string? RpcProxyServer { get; set; }

    [JsonPropertyName("exchangeServer")]
    public string? ExchangeServer { get; set; }

    [JsonPropertyName("emailAddress")]
    public string? EmailAddress { get; set; }

    [JsonPropertyName("remoteTenant")]
    public string? RemoteTenant { get; set; }

    [JsonPropertyName("port")]
    public int? Port { get; set; }

    [JsonPropertyName("security")]
    public string? Security { get; set; }

    [JsonPropertyName("authentication")]
    public string? Authentication { get; set; }

    [JsonPropertyName("maxConcurrentMigrations")]
    public int? MaxConcurrentMigrations { get; set; }

    [JsonPropertyName("maxConcurrentIncrementalSyncs")]
    public int? MaxConcurrentIncrementalSyncs { get; set; }

    [JsonPropertyName("skipVerification")]
    public bool? SkipVerification { get; set; }

    [JsonPropertyName("acceptUntrustedCertificates")]
    public bool? AcceptUntrustedCertificates { get; set; }

    [JsonPropertyName("lastModifiedTime")]
    public DateTime? LastModifiedTime { get; set; }
}

public class GetMigrationEndpointsRequest
{
    [JsonPropertyName("searchQuery")]
    public string? SearchQuery { get; set; }

    [JsonPropertyName("sortBy")]
    public string? SortBy { get; set; }

    [JsonPropertyName("sortDescending")]
    public bool SortDescending { get; set; }
}

public class GetMigrationEndpointsResponse
{
    [JsonPropertyName("endpoints")]
    public List<MigrationEndpointDto> Endpoints { get; set; } = new();
}

public class UpsertMigrationEndpointRequest
{
    [JsonPropertyName("identity")]
    public string? Identity { get; set; }

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("endpointType")]
    public string EndpointType { get; set; } = string.Empty;

    [JsonPropertyName("remoteServer")]
    public string? RemoteServer { get; set; }

    [JsonPropertyName("rpcProxyServer")]
    public string? RpcProxyServer { get; set; }

    [JsonPropertyName("exchangeServer")]
    public string? ExchangeServer { get; set; }

    [JsonPropertyName("emailAddress")]
    public string? EmailAddress { get; set; }

    [JsonPropertyName("remoteTenant")]
    public string? RemoteTenant { get; set; }

    [JsonPropertyName("port")]
    public int? Port { get; set; }

    [JsonPropertyName("security")]
    public string? Security { get; set; }

    [JsonPropertyName("authentication")]
    public string? Authentication { get; set; }

    [JsonPropertyName("username")]
    public string? Username { get; set; }

    [JsonIgnore]
    public string? Password { get; set; }

    [JsonPropertyName("passwordSecret")]
    public ProtectedSecretReference? PasswordSecret { get; set; }

    [JsonPropertyName("maxConcurrentMigrations")]
    public int? MaxConcurrentMigrations { get; set; }

    [JsonPropertyName("maxConcurrentIncrementalSyncs")]
    public int? MaxConcurrentIncrementalSyncs { get; set; }

    [JsonPropertyName("skipVerification")]
    public bool SkipVerification { get; set; }

    [JsonPropertyName("acceptUntrustedCertificates")]
    public bool AcceptUntrustedCertificates { get; set; }
}

public class TestMigrationEndpointRequest
{
    [JsonPropertyName("identity")]
    public string? Identity { get; set; }

    [JsonPropertyName("useExistingEndpoint")]
    public bool UseExistingEndpoint { get; set; }

    [JsonPropertyName("endpointType")]
    public string EndpointType { get; set; } = string.Empty;

    [JsonPropertyName("remoteServer")]
    public string? RemoteServer { get; set; }

    [JsonPropertyName("rpcProxyServer")]
    public string? RpcProxyServer { get; set; }

    [JsonPropertyName("exchangeServer")]
    public string? ExchangeServer { get; set; }

    [JsonPropertyName("emailAddress")]
    public string? EmailAddress { get; set; }

    [JsonPropertyName("port")]
    public int? Port { get; set; }

    [JsonPropertyName("security")]
    public string? Security { get; set; }

    [JsonPropertyName("authentication")]
    public string? Authentication { get; set; }

    [JsonPropertyName("username")]
    public string? Username { get; set; }

    [JsonIgnore]
    public string? Password { get; set; }

    [JsonPropertyName("passwordSecret")]
    public ProtectedSecretReference? PasswordSecret { get; set; }

    [JsonPropertyName("skipVerification")]
    public bool SkipVerification { get; set; }

    [JsonPropertyName("acceptUntrustedCertificates")]
    public bool AcceptUntrustedCertificates { get; set; }
}

public class TestMigrationEndpointResponse
{
    [JsonPropertyName("summary")]
    public string Summary { get; set; } = string.Empty;

    [JsonPropertyName("details")]
    public string? Details { get; set; }
}

public class GetMigrationBatchPreflightRequest
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("batchType")]
    public string BatchType { get; set; } = string.Empty;

    [JsonPropertyName("endpointIdentity")]
    public string EndpointIdentity { get; set; } = string.Empty;

    [JsonPropertyName("csvFilePath")]
    public string CsvFilePath { get; set; } = string.Empty;

    [JsonPropertyName("targetDeliveryDomain")]
    public string? TargetDeliveryDomain { get; set; }
}

public class GetMigrationBatchPreflightResponse
{
    [JsonPropertyName("isReady")]
    public bool IsReady { get; set; }

    [JsonPropertyName("endpointType")]
    public string? EndpointType { get; set; }

    [JsonPropertyName("csvRowCount")]
    public int CsvRowCount { get; set; }

    [JsonPropertyName("csvHeaders")]
    public List<string> CsvHeaders { get; set; } = new();

    [JsonPropertyName("messages")]
    public List<string> Messages { get; set; } = new();
}

public class CreateMigrationBatchRequest
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("batchType")]
    public string BatchType { get; set; } = string.Empty;

    [JsonPropertyName("endpointIdentity")]
    public string EndpointIdentity { get; set; } = string.Empty;

    [JsonPropertyName("csvFilePath")]
    public string CsvFilePath { get; set; } = string.Empty;

    [JsonPropertyName("targetDeliveryDomain")]
    public string? TargetDeliveryDomain { get; set; }

    [JsonPropertyName("notificationEmails")]
    public List<string> NotificationEmails { get; set; } = new();

    [JsonPropertyName("autoStart")]
    public bool AutoStart { get; set; }

    [JsonPropertyName("autoComplete")]
    public bool AutoComplete { get; set; }
}

public class GetMigrationBatchesRequest
{
    [JsonPropertyName("searchQuery")]
    public string? SearchQuery { get; set; }

    [JsonPropertyName("status")]
    public string? Status { get; set; }

    [JsonPropertyName("pageSize")]
    public int PageSize { get; set; } = PagingDefaults.DefaultPageSize;

    [JsonPropertyName("skip")]
    public int Skip { get; set; }

    [JsonPropertyName("sortBy")]
    public string? SortBy { get; set; }

    [JsonPropertyName("sortDescending")]
    public bool SortDescending { get; set; }
}

public class GetMigrationBatchesResponse
{
    [JsonPropertyName("batches")]
    public List<MigrationBatchListItemDto> Batches { get; set; } = new();

    [JsonPropertyName("totalCount")]
    public int TotalCount { get; set; }

    [JsonPropertyName("skip")]
    public int Skip { get; set; }

    [JsonPropertyName("pageSize")]
    public int PageSize { get; set; }

    [JsonPropertyName("hasMore")]
    public bool HasMore { get; set; }

    [JsonPropertyName("searchQuery")]
    public string? SearchQuery { get; set; }
}

public class GetMigrationBatchDetailsRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;
}

public class StartMigrationBatchRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;
}

public class CompleteMigrationBatchRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;
}

public class RemoveMigrationBatchRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;
}
