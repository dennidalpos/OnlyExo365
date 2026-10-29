using System.Collections;
using OnlyExo365.Contracts.Dtos;

namespace OnlyExo365.Worker.PowerShell;

internal static class ExoMailboxMapper
{
    internal static MailboxListItemDto ToMailboxListItem(Hashtable hash)
    {
        return new MailboxListItemDto
        {
            Identity = hash["Identity"]?.ToString() ?? string.Empty,
            Guid = hash["Guid"]?.ToString(),
            DisplayName = hash["DisplayName"]?.ToString() ?? string.Empty,
            PrimarySmtpAddress = hash["PrimarySmtpAddress"]?.ToString() ?? string.Empty,
            UserPrincipalName = hash["UserPrincipalName"]?.ToString(),
            RecipientType = hash["RecipientType"]?.ToString() ?? string.Empty,
            RecipientTypeDetails = hash["RecipientTypeDetails"]?.ToString() ?? string.Empty,
            Alias = hash["Alias"]?.ToString(),
            IsInactiveMailbox = hash["IsInactiveMailbox"] as bool? ?? false
        };
    }

    internal static DeletedMailboxItemDto ToDeletedMailboxItem(Hashtable hash)
    {
        return new DeletedMailboxItemDto
        {
            Identity = hash["Identity"]?.ToString() ?? string.Empty,
            DisplayName = hash["DisplayName"]?.ToString() ?? string.Empty,
            UserPrincipalName = hash["UserPrincipalName"]?.ToString(),
            PrimarySmtpAddress = hash["PrimarySmtpAddress"]?.ToString() ?? string.Empty,
            RecipientTypeDetails = hash["RecipientTypeDetails"]?.ToString() ?? string.Empty,
            Alias = hash["Alias"]?.ToString(),
            DeletionType = ParseEnum(hash["DeletionType"], DeletedMailboxDeletionType.Unknown)
        };
    }

    internal static MailboxDetailsDto ToMailboxDetails(Hashtable hash)
    {
        return new MailboxDetailsDto
        {
            Identity = hash["Identity"]?.ToString() ?? string.Empty,
            Guid = hash["Guid"]?.ToString(),
            DisplayName = hash["DisplayName"]?.ToString() ?? string.Empty,
            PrimarySmtpAddress = hash["PrimarySmtpAddress"]?.ToString() ?? string.Empty,
            UserPrincipalName = hash["UserPrincipalName"]?.ToString(),
            Alias = hash["Alias"]?.ToString(),
            RecipientType = hash["RecipientType"]?.ToString() ?? string.Empty,
            RecipientTypeDetails = hash["RecipientTypeDetails"]?.ToString() ?? string.Empty,
            EmailAddresses = ConvertToStringList(hash["EmailAddresses"]),
            RetentionPolicy = hash["RetentionPolicy"]?.ToString(),
            WhenCreated = hash["WhenCreated"] as DateTime?,
            WhenMailboxCreated = hash["WhenMailboxCreated"] as DateTime?,
            Features = new MailboxFeaturesDto
            {
                ArchiveEnabled = hash["ArchiveEnabled"] as bool? ?? false,
                ArchiveName = hash["ArchiveName"]?.ToString(),
                ArchiveGuid = hash["ArchiveGuid"]?.ToString(),
                ArchiveStatus = hash["ArchiveStatus"]?.ToString(),
                LitigationHoldEnabled = hash["LitigationHoldEnabled"] as bool? ?? false,
                LitigationHoldDate = hash["LitigationHoldDate"] as DateTime?,
                LitigationHoldOwner = hash["LitigationHoldOwner"]?.ToString(),
                LitigationHoldDuration = hash["LitigationHoldDuration"]?.ToString(),
                AuditEnabled = hash["AuditEnabled"] as bool? ?? false,
                AuditLogAgeLimit = hash["AuditLogAgeLimit"]?.ToString(),
                AuditAdmin = ConvertToStringList(hash["AuditAdmin"]),
                AuditDelegate = ConvertToStringList(hash["AuditDelegate"]),
                AuditOwner = ConvertToStringList(hash["AuditOwner"]),
                ForwardingAddress = hash["ForwardingAddress"]?.ToString(),
                ForwardingSmtpAddress = hash["ForwardingSmtpAddress"]?.ToString(),
                DeliverToMailboxAndForward = hash["DeliverToMailboxAndForward"] as bool? ?? false,
                ProhibitSendQuota = hash["ProhibitSendQuota"]?.ToString(),
                ProhibitSendQuotaBytes = hash["ProhibitSendQuotaBytes"] as long?,
                ProhibitSendReceiveQuota = hash["ProhibitSendReceiveQuota"]?.ToString(),
                ProhibitSendReceiveQuotaBytes = hash["ProhibitSendReceiveQuotaBytes"] as long?,
                IssueWarningQuota = hash["IssueWarningQuota"]?.ToString(),
                IssueWarningQuotaBytes = hash["IssueWarningQuotaBytes"] as long?,
                MaxSendSize = hash["MaxSendSize"]?.ToString(),
                MaxReceiveSize = hash["MaxReceiveSize"]?.ToString(),
                RetentionHoldEnabled = hash["RetentionHoldEnabled"] as bool? ?? false,
                SingleItemRecoveryEnabled = hash["SingleItemRecoveryEnabled"] as bool? ?? false,
                RetainDeletedItemsFor = hash["RetainDeletedItemsFor"]?.ToString(),
                HiddenFromAddressListsEnabled = hash["HiddenFromAddressListsEnabled"] as bool? ?? false
            }
        };
    }

    internal static void ApplyCasMailboxSettings(MailboxFeaturesDto features, Hashtable hash)
    {
        features.OwaEnabled = hash["OwaEnabled"] as bool?;
        features.ActiveSyncEnabled = hash["ActiveSyncEnabled"] as bool?;
        features.MapiEnabled = hash["MapiEnabled"] as bool?;
        features.PopEnabled = hash["PopEnabled"] as bool?;
        features.ImapEnabled = hash["ImapEnabled"] as bool?;
        features.SmtpClientAuthenticationDisabled = hash["SmtpClientAuthenticationDisabled"] as bool?;
    }

    internal static MailboxStatisticsDto ToMailboxStatistics(Hashtable hash)
    {
        return new MailboxStatisticsDto
        {
            TotalItemSize = hash["TotalItemSize"]?.ToString(),
            TotalItemSizeBytes = hash["TotalItemSizeBytes"] as long?,
            ItemCount = Convert.ToInt32(hash["ItemCount"] ?? 0),
            DeletedItemCount = Convert.ToInt32(hash["DeletedItemCount"] ?? 0),
            TotalDeletedItemSize = hash["TotalDeletedItemSize"]?.ToString(),
            LastLogonTime = hash["LastLogonTime"] as DateTime?,
            LastLogoffTime = hash["LastLogoffTime"] as DateTime?
        };
    }

    internal static InboxRuleDto ToInboxRule(Hashtable hash)
    {
        return new InboxRuleDto
        {
            Name = hash["Name"]?.ToString() ?? string.Empty,
            RuleIdentity = hash["RuleIdentity"]?.ToString(),
            Enabled = hash["Enabled"] as bool? ?? false,
            Priority = Convert.ToInt32(hash["Priority"] ?? 0),
            Description = hash["Description"]?.ToString(),
            ForwardTo = ConvertToStringList(hash["ForwardTo"]),
            ForwardAsAttachmentTo = ConvertToStringList(hash["ForwardAsAttachmentTo"]),
            RedirectTo = ConvertToStringList(hash["RedirectTo"]),
            DeleteMessage = hash["DeleteMessage"] as bool? ?? false,
            MoveToFolder = hash["MoveToFolder"]?.ToString()
        };
    }

    internal static AutoReplyConfigurationDto ToAutoReplyConfiguration(Hashtable hash)
    {
        return new AutoReplyConfigurationDto
        {
            AutoReplyState = hash["AutoReplyState"]?.ToString() ?? "Disabled",
            StartTime = hash["StartTime"] as DateTime?,
            EndTime = hash["EndTime"] as DateTime?,
            InternalMessage = hash["InternalMessage"]?.ToString(),
            ExternalMessage = hash["ExternalMessage"]?.ToString(),
            ExternalAudience = hash["ExternalAudience"]?.ToString()
        };
    }

    internal static RetentionPolicySummaryDto ToRetentionPolicySummary(Hashtable hash)
    {
        return new RetentionPolicySummaryDto
        {
            Id = hash["Id"]?.ToString(),
            Name = hash["Name"]?.ToString() ?? string.Empty,
            Description = hash["Description"]?.ToString(),
            RequiresArchive = hash["RequiresArchive"] as bool? ?? false
        };
    }

    internal static RestoreMailboxResponse ToRestoreMailboxResponse(PowerShellResult result, RestoreMailboxRequest request)
    {
        var response = new RestoreMailboxResponse
        {
            SourceIdentity = request.SourceIdentity,
            TargetMailbox = request.TargetMailbox,
            Scenario = RestoreMailboxScenario.Unknown,
            Status = RestoreMailboxStatus.NotStarted
        };

        if (result.Output.Count > 0 && result.Output.First().BaseObject is Hashtable hash)
        {
            var scenarioValue = hash["Scenario"]?.ToString();
            if (!string.IsNullOrWhiteSpace(scenarioValue) &&
                Enum.TryParse(scenarioValue, true, out RestoreMailboxScenario parsedScenario))
            {
                response.Scenario = parsedScenario;
            }

            response.Action = hash["Action"]?.ToString();

            var statusValue = hash["Status"]?.ToString();
            if (!string.IsNullOrWhiteSpace(statusValue))
            {
                response.Status = MapRestoreStatus(statusValue);
            }

            response.StatusDetail = hash["StatusDetail"]?.ToString();

            if (hash["PercentComplete"] != null &&
                int.TryParse(hash["PercentComplete"]?.ToString(), out var percent))
            {
                response.PercentComplete = percent;
            }

            response.RequestGuid = hash["RequestGuid"]?.ToString();

            var errorCodeValue = hash["ErrorCode"]?.ToString();
            if (!string.IsNullOrWhiteSpace(errorCodeValue) &&
                Enum.TryParse(errorCodeValue, true, out RestoreMailboxErrorCode parsedErrorCode) &&
                parsedErrorCode != RestoreMailboxErrorCode.None)
            {
                response.Error = new RestoreMailboxErrorDto
                {
                    Code = parsedErrorCode,
                    Message = hash["ErrorMessage"]?.ToString() ?? "Mailbox restore failed"
                };
            }

            if (response.Error == null && response.Status == RestoreMailboxStatus.Failed)
            {
                response.Error = new RestoreMailboxErrorDto
                {
                    Code = RestoreMailboxErrorCode.Unknown,
                    Message = hash["ErrorMessage"]?.ToString() ?? "Mailbox restore failed"
                };
            }

            response.SourceIdentity = hash["SourceIdentity"]?.ToString() ?? response.SourceIdentity;
            response.TargetMailbox = hash["TargetMailbox"]?.ToString() ?? response.TargetMailbox;
        }
        else if (!result.Success)
        {
            response.Status = RestoreMailboxStatus.Failed;
            response.Error = new RestoreMailboxErrorDto
            {
                Code = RestoreMailboxErrorCode.Unknown,
                Message = result.ErrorMessage ?? "Mailbox restore failed"
            };
        }

        return response;
    }

    internal static List<string> ConvertToStringList(object? obj)
    {
        if (obj == null)
        {
            return new List<string>();
        }

        if (obj is object[] array)
        {
            return array
                .Select(x => x?.ToString() ?? string.Empty)
                .Where(x => !string.IsNullOrEmpty(x))
                .ToList();
        }

        if (obj is IEnumerable enumerable)
        {
            var list = new List<string>();
            foreach (var item in enumerable)
            {
                var str = item?.ToString();
                if (!string.IsNullOrEmpty(str))
                {
                    list.Add(str);
                }
            }

            return list;
        }

        var single = obj.ToString();
        return string.IsNullOrEmpty(single)
            ? new List<string>()
            : new List<string> { single };
    }

    private static RestoreMailboxStatus MapRestoreStatus(string value)
    {
        if (Enum.TryParse(value, true, out RestoreMailboxStatus parsed))
        {
            return parsed;
        }

        return value.ToLowerInvariant() switch
        {
            "queued" => RestoreMailboxStatus.InProgress,
            "inprogress" => RestoreMailboxStatus.InProgress,
            "completed" => RestoreMailboxStatus.Completed,
            "failed" => RestoreMailboxStatus.Failed,
            "suspended" => RestoreMailboxStatus.Failed,
            _ => RestoreMailboxStatus.NotStarted
        };
    }

    private static TEnum ParseEnum<TEnum>(object? value, TEnum fallback)
        where TEnum : struct
    {
        var raw = value?.ToString();
        return !string.IsNullOrWhiteSpace(raw) && Enum.TryParse(raw, true, out TEnum parsed)
            ? parsed
            : fallback;
    }
}

