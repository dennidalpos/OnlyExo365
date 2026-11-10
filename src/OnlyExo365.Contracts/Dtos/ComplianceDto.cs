using System.Text.Json.Serialization;

namespace OnlyExo365.Contracts.Dtos;

public class GetComplianceWorkspaceRequest
{
    [JsonPropertyName("maxActions")]
    public int MaxActions { get; set; } = 50;
}

public class GetComplianceWorkspaceResponse
{
    [JsonPropertyName("correlationId")]
    public string? CorrelationId { get; set; }

    [JsonPropertyName("searches")]
    public List<ComplianceSearchDto> Searches { get; set; } = new();

    [JsonPropertyName("cases")]
    public List<ComplianceCaseDto> Cases { get; set; } = new();

    [JsonPropertyName("actions")]
    public List<ComplianceActionSummaryDto> Actions { get; set; } = new();

    [JsonPropertyName("warnings")]
    public List<string> Warnings { get; set; } = new();

    [JsonPropertyName("warningDetails")]
    public List<OperationWarningDto> WarningDetails { get; set; } = new();

    [JsonPropertyName("hasPartialData")]
    public bool HasPartialData { get; set; }

    [JsonPropertyName("isHoldListingUnsupported")]
    public bool IsHoldListingUnsupported { get; set; }

    [JsonPropertyName("holdListingStatusMessage")]
    public string? HoldListingStatusMessage { get; set; }
}

public class ComplianceSearchDto
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("caseName")]
    public string? CaseName { get; set; }

    [JsonPropertyName("status")]
    public string? Status { get; set; }

    [JsonPropertyName("createdBy")]
    public string? CreatedBy { get; set; }

    [JsonPropertyName("createdTime")]
    public DateTime? CreatedTime { get; set; }

    [JsonPropertyName("lastModifiedTime")]
    public DateTime? LastModifiedTime { get; set; }

    [JsonPropertyName("items")]
    public string? Items { get; set; }

    [JsonPropertyName("size")]
    public string? Size { get; set; }

    [JsonPropertyName("exchangeLocations")]
    public List<string> ExchangeLocations { get; set; } = new();

    [JsonPropertyName("contentMatchQuery")]
    public string? ContentMatchQuery { get; set; }
}

public class ComplianceCaseDto
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("status")]
    public string? Status { get; set; }

    [JsonPropertyName("caseType")]
    public string? CaseType { get; set; }

    [JsonPropertyName("createdTime")]
    public DateTime? CreatedTime { get; set; }

    [JsonPropertyName("description")]
    public string? Description { get; set; }
}

public class ComplianceActionSummaryDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("actionType")]
    public string ActionType { get; set; } = string.Empty;

    [JsonPropertyName("searchName")]
    public string? SearchName { get; set; }

    [JsonPropertyName("caseName")]
    public string? CaseName { get; set; }

    [JsonPropertyName("status")]
    public string? Status { get; set; }

    [JsonPropertyName("createdBy")]
    public string? CreatedBy { get; set; }

    [JsonPropertyName("createdTime")]
    public DateTime? CreatedTime { get; set; }

    [JsonPropertyName("completedTime")]
    public DateTime? CompletedTime { get; set; }

    [JsonPropertyName("exchangeLocations")]
    public List<string> ExchangeLocations { get; set; } = new();

    [JsonPropertyName("details")]
    public string? Details { get; set; }
}

public class CreateComplianceSearchRequest
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("caseName")]
    public string? CaseName { get; set; }

    [JsonPropertyName("exchangeLocations")]
    public List<string> ExchangeLocations { get; set; } = new();

    [JsonPropertyName("contentMatchQuery")]
    public string? ContentMatchQuery { get; set; }
}

public class StartComplianceSearchRequest
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;
}

public class RemoveComplianceSearchRequest
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;
}

public class InvokeComplianceActionRequest
{
    [JsonPropertyName("searchName")]
    public string SearchName { get; set; } = string.Empty;

    [JsonPropertyName("actionType")]
    public string ActionType { get; set; } = string.Empty;

    [JsonPropertyName("purgeType")]
    public string? PurgeType { get; set; }

    [JsonPropertyName("caseName")]
    public string? CaseName { get; set; }

    [JsonPropertyName("holdName")]
    public string? HoldName { get; set; }
}

public class InvokeComplianceActionResponse
{
    [JsonPropertyName("action")]
    public ComplianceActionSummaryDto Action { get; set; } = new();
}

public class SearchUnifiedAuditLogRequest
{
    [JsonPropertyName("startDate")]
    public DateTime StartDate { get; set; }

    [JsonPropertyName("endDate")]
    public DateTime EndDate { get; set; }

    [JsonPropertyName("userIds")]
    public List<string> UserIds { get; set; } = new();

    [JsonPropertyName("operations")]
    public List<string> Operations { get; set; } = new();

    [JsonPropertyName("objectIds")]
    public List<string> ObjectIds { get; set; } = new();

    [JsonPropertyName("freeText")]
    public string? FreeText { get; set; }

    [JsonPropertyName("maxResults")]
    public int MaxResults { get; set; } = 100;
}

public class SearchUnifiedAuditLogResponse
{
    [JsonPropertyName("results")]
    public List<UnifiedAuditLogRecordDto> Results { get; set; } = new();

    [JsonPropertyName("totalCount")]
    public int TotalCount { get; set; }

    [JsonPropertyName("warning")]
    public string? Warning { get; set; }
}

public class UnifiedAuditLogRecordDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("creationDate")]
    public DateTime? CreationDate { get; set; }

    [JsonPropertyName("userIds")]
    public string? UserIds { get; set; }

    [JsonPropertyName("operations")]
    public string? Operations { get; set; }

    [JsonPropertyName("recordType")]
    public string? RecordType { get; set; }

    [JsonPropertyName("resultStatus")]
    public string? ResultStatus { get; set; }

    [JsonPropertyName("objectId")]
    public string? ObjectId { get; set; }

    [JsonPropertyName("auditData")]
    public string? AuditData { get; set; }
}

