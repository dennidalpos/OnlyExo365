using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Dtos;

public class TenantLicenseDto
{
    [JsonPropertyName("skuId")]
    public string SkuId { get; set; } = string.Empty;

    [JsonPropertyName("skuPartNumber")]
    public string SkuPartNumber { get; set; } = string.Empty;

    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;

    [JsonPropertyName("total")]
    public int Total { get; set; }

    [JsonPropertyName("assigned")]
    public int Assigned { get; set; }

    [JsonPropertyName("available")]
    public int Available { get; set; }
}

public class AdminRoleMemberDto
{
    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;

    [JsonPropertyName("userPrincipalName")]
    public string UserPrincipalName { get; set; } = string.Empty;

    [JsonPropertyName("roleName")]
    public string RoleName { get; set; } = string.Empty;
}

public class UserLicenseDto
{
    [JsonPropertyName("skuId")]
    public string SkuId { get; set; } = string.Empty;

    [JsonPropertyName("skuPartNumber")]
    public string SkuPartNumber { get; set; } = string.Empty;

    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;
}

public class GetAvailableLicensesResponse
{
    [JsonPropertyName("licenses")]
    public List<TenantLicenseDto> Licenses { get; set; } = new();
}

public class SetUserLicenseRequest
{
    [JsonPropertyName("userPrincipalName")]
    public string UserPrincipalName { get; set; } = string.Empty;

    [JsonPropertyName("addLicenseSkuIds")]
    public List<string> AddLicenseSkuIds { get; set; } = new();

    [JsonPropertyName("removeLicenseSkuIds")]
    public List<string> RemoveLicenseSkuIds { get; set; } = new();
}

public class GetUserLicensesRequest
{
    [JsonPropertyName("userPrincipalName")]
    public string UserPrincipalName { get; set; } = string.Empty;
}

public class GetUserLicensesResponse
{
    [JsonPropertyName("licenses")]
    public List<UserLicenseDto> Licenses { get; set; } = new();
}

public class PrerequisiteStatusDto
{
    [JsonPropertyName("powerShellVersion")]
    public string? PowerShellVersion { get; set; }

    [JsonPropertyName("isPowerShell7")]
    public bool IsPowerShell7 { get; set; }

    [JsonPropertyName("exchangeModuleInstalled")]
    public bool ExchangeModuleInstalled { get; set; }

    [JsonPropertyName("exchangeModuleVersion")]
    public string? ExchangeModuleVersion { get; set; }

    [JsonPropertyName("graphModuleInstalled")]
    public bool GraphModuleInstalled { get; set; }

    [JsonPropertyName("graphModuleVersion")]
    public string? GraphModuleVersion { get; set; }
}

public class InstallModuleRequest
{
    [JsonPropertyName("moduleName")]
    public string ModuleName { get; set; } = string.Empty;
}

public class InstallModuleResponse
{
    [JsonPropertyName("success")]
    public bool Success { get; set; }

    [JsonPropertyName("message")]
    public string Message { get; set; } = string.Empty;

    [JsonPropertyName("moduleName")]
    public string ModuleName { get; set; } = string.Empty;

    [JsonPropertyName("installedVersion")]
    public string? InstalledVersion { get; set; }

    [JsonPropertyName("manualInstructions")]
    public string? ManualInstructions { get; set; }
}
