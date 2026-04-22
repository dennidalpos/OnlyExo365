using OnlyExo365.Contracts.Dtos;

namespace OnlyExo365.Worker.PowerShell;

internal sealed class ExoOrganizationRelationshipCommands : ExoCommandModuleBase
{
    public ExoOrganizationRelationshipCommands(PowerShellEngine engine)
        : base(engine)
    {
    }

    public async Task<GetOrganizationRelationshipsResponse> GetOrganizationRelationshipsAsync(CancellationToken cancellationToken = default)
    {
        var script = @"
Get-OrganizationRelationship -ErrorAction Stop |
    Sort-Object Name |
    ForEach-Object {
        [PSCustomObject]@{
            Identity = $_.Identity.ToString()
            Name = $_.Name
            DomainNames = @($_.DomainNames | ForEach-Object { $_.ToString() })
            Enabled = [bool]$_.Enabled
            FreeBusyAccessEnabled = [bool]$_.FreeBusyAccessEnabled
            FreeBusyAccessLevel = if ($_.FreeBusyAccessLevel) { $_.FreeBusyAccessLevel.ToString() } else { 'AvailabilityOnly' }
            MailTipsAccessEnabled = [bool]$_.MailTipsAccessEnabled
            MailTipsAccessLevel = if ($_.MailTipsAccessLevel) { $_.MailTipsAccessLevel.ToString() } else { 'All' }
            TargetApplicationUri = if ($_.TargetApplicationUri) { $_.TargetApplicationUri.ToString() } else { $null }
            TargetAutodiscoverEpr = if ($_.TargetAutodiscoverEpr) { $_.TargetAutodiscoverEpr.ToString() } else { $null }
            ArchiveAccessEnabled = if ($null -ne $_.ArchiveAccessEnabled) { [bool]$_.ArchiveAccessEnabled } else { $null }
            DeliveryReportEnabled = if ($null -ne $_.DeliveryReportEnabled) { [bool]$_.DeliveryReportEnabled } else { $null }
            MailboxMoveEnabled = if ($null -ne $_.MailboxMoveEnabled) { [bool]$_.MailboxMoveEnabled } else { $null }
            PhotosEnabled = if ($null -ne $_.PhotosEnabled) { [bool]$_.PhotosEnabled } else { $null }
        }
    }";

        var results = await RunScriptAsync(script, cancellationToken);
        var relationships = new List<OrganizationRelationshipDto>();

        foreach (var obj in results)
        {
            relationships.Add(new OrganizationRelationshipDto
            {
                Identity = GetString(obj, "Identity"),
                Name = GetString(obj, "Name"),
                DomainNames = ConvertToStringList(obj.Properties["DomainNames"]?.Value),
                Enabled = GetBool(obj, "Enabled"),
                FreeBusyAccessEnabled = GetBool(obj, "FreeBusyAccessEnabled"),
                FreeBusyAccessLevel = GetString(obj, "FreeBusyAccessLevel"),
                MailTipsAccessEnabled = GetBool(obj, "MailTipsAccessEnabled"),
                MailTipsAccessLevel = GetString(obj, "MailTipsAccessLevel"),
                TargetApplicationUri = GetNullableString(obj, "TargetApplicationUri"),
                TargetAutodiscoverEpr = GetNullableString(obj, "TargetAutodiscoverEpr"),
                ArchiveAccessEnabled = GetNullableBool(obj, "ArchiveAccessEnabled"),
                DeliveryReportEnabled = GetNullableBool(obj, "DeliveryReportEnabled"),
                MailboxMoveEnabled = GetNullableBool(obj, "MailboxMoveEnabled"),
                PhotosEnabled = GetNullableBool(obj, "PhotosEnabled")
            });
        }

        return new GetOrganizationRelationshipsResponse
        {
            Relationships = relationships
        };
    }

    public async Task UpsertOrganizationRelationshipAsync(UpsertOrganizationRelationshipRequest request, CancellationToken cancellationToken = default)
    {
        request.FreeBusyAccessLevel = NormalizeFreeBusyAccessLevel(request.FreeBusyAccessLevel);
        request.MailTipsAccessLevel = NormalizeMailTipsAccessLevel(request.MailTipsAccessLevel);
        var script = OrganizationRelationshipCommandBuilder.BuildUpsertOrganizationRelationshipScript(request);
        await RunScriptAsync(script, cancellationToken);
    }

    public async Task RemoveOrganizationRelationshipAsync(RemoveOrganizationRelationshipRequest request, CancellationToken cancellationToken = default)
    {
        var identity = EscapePs(request.Identity);
        var script = $@"
Remove-OrganizationRelationship -Identity '{identity}' -Confirm:$false -ErrorAction Stop
Write-Output 'OK'";

        await RunScriptAsync(script, cancellationToken);
    }

    private static string NormalizeFreeBusyAccessLevel(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "AvailabilityOnly";
        }

        return value.Trim() switch
        {
            "None" => "None",
            "LimitedDetails" => "LimitedDetails",
            "AvailabilityOnly" => "AvailabilityOnly",
            _ => "AvailabilityOnly"
        };
    }

    private static string NormalizeMailTipsAccessLevel(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "All";
        }

        return value.Trim() switch
        {
            "None" => "None",
            "Limited" => "Limited",
            "All" => "All",
            _ => "All"
        };
    }
}

