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

    [JsonPropertyName("from")]
    public List<string> From { get; set; } = new();

    [JsonPropertyName("sentTo")]
    public List<string> SentTo { get; set; } = new();

    [JsonPropertyName("subjectContainsWords")]
    public List<string> SubjectContainsWords { get; set; } = new();

    [JsonPropertyName("prependSubject")]
    public string PrependSubject { get; set; } = string.Empty;
}

public class SetTransportRuleStateRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; }
}

public class RemoveTransportRuleRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;
}

public class UpsertTransportRuleRequest
{
    [JsonPropertyName("identity")]
    public string? Identity { get; set; }

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("from")]
    public List<string> From { get; set; } = new();

    [JsonPropertyName("sentTo")]
    public List<string> SentTo { get; set; } = new();

    [JsonPropertyName("subjectContainsWords")]
    public List<string> SubjectContainsWords { get; set; } = new();

    [JsonPropertyName("prependSubject")]
    public string? PrependSubject { get; set; }

    [JsonPropertyName("mode")]
    public string? Mode { get; set; }

    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; } = true;
}

public class TestTransportRuleRequest
{
    [JsonPropertyName("sender")]
    public string Sender { get; set; } = string.Empty;

    [JsonPropertyName("recipient")]
    public string Recipient { get; set; } = string.Empty;

    [JsonPropertyName("subject")]
    public string Subject { get; set; } = string.Empty;
}

public class TestTransportRuleResponse
{
    [JsonPropertyName("matchedRuleNames")]
    public List<string> MatchedRuleNames { get; set; } = new();
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
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("displayLabel")]
    public string DisplayLabel { get; set; } = string.Empty;

    [JsonPropertyName("type")]
    public string Type { get; set; } = string.Empty;

    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; }

    [JsonPropertyName("comment")]
    public string Comment { get; set; } = string.Empty;

    [JsonPropertyName("senderDomains")]
    public List<string> SenderDomains { get; set; } = new();

    [JsonPropertyName("recipientDomains")]
    public List<string> RecipientDomains { get; set; } = new();
}

public class UpsertConnectorRequest
{
    [JsonPropertyName("identity")]
    public string? Identity { get; set; }

    [JsonPropertyName("type")]
    public string Type { get; set; } = "Inbound";

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; } = true;

    [JsonPropertyName("comment")]
    public string? Comment { get; set; }

    [JsonPropertyName("senderDomains")]
    public List<string> SenderDomains { get; set; } = new();

    [JsonPropertyName("recipientDomains")]
    public List<string> RecipientDomains { get; set; } = new();
}

public class RemoveConnectorRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("type")]
    public string Type { get; set; } = string.Empty;
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
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("domainName")]
    public string DomainName { get; set; } = string.Empty;

    [JsonPropertyName("domainType")]
    public string DomainType { get; set; } = string.Empty;

    [JsonPropertyName("default")]
    public bool Default { get; set; }
}

public class UpsertAcceptedDomainRequest
{
    [JsonPropertyName("identity")]
    public string? Identity { get; set; }

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("domainName")]
    public string DomainName { get; set; } = string.Empty;

    [JsonPropertyName("domainType")]
    public string DomainType { get; set; } = "Authoritative";

    [JsonPropertyName("makeDefault")]
    public bool MakeDefault { get; set; }
}

public class RemoveAcceptedDomainRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;
}
