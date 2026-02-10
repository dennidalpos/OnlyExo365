using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Dtos;

public class GetTransportRulesRequest
{
}

public class GetTransportRulesResponse
{
    [JsonPropertyName("rules")]
    public List<TransportRuleDto> Rules { get; set; } = new();
}

public class TransportRuleDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("priority")]
    public int? Priority { get; set; }

    [JsonPropertyName("state")]
    public string State { get; set; } = string.Empty;

    [JsonPropertyName("mode")]
    public string Mode { get; set; } = string.Empty;

    [JsonPropertyName("description")]
    public string Description { get; set; } = string.Empty;
}

public class SetTransportRuleStateRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; }
}

public class GetConnectorsRequest
{
}

public class GetConnectorsResponse
{
    [JsonPropertyName("connectors")]
    public List<ConnectorDto> Connectors { get; set; } = new();
}

public class ConnectorDto
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("type")]
    public string Type { get; set; } = string.Empty;

    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; }

    [JsonPropertyName("comment")]
    public string Comment { get; set; } = string.Empty;
}

public class GetAcceptedDomainsRequest
{
}

public class GetAcceptedDomainsResponse
{
    [JsonPropertyName("domains")]
    public List<AcceptedDomainDto> Domains { get; set; } = new();
}

public class AcceptedDomainDto
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("domainName")]
    public string DomainName { get; set; } = string.Empty;

    [JsonPropertyName("domainType")]
    public string DomainType { get; set; } = string.Empty;

    [JsonPropertyName("default")]
    public bool Default { get; set; }
}
