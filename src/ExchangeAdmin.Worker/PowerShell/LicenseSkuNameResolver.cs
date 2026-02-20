using System.Globalization;

namespace ExchangeAdmin.Worker.PowerShell;

internal static class LicenseSkuNameResolver
{
    private static readonly Dictionary<string, string> FriendlyNames = new(StringComparer.OrdinalIgnoreCase)
    {
        ["O365_BUSINESS_ESSENTIALS"] = "Microsoft 365 Business Basic",
        ["O365_BUSINESS_PREMIUM"] = "Microsoft 365 Business Standard",
        ["SPB"] = "Microsoft 365 Business Premium",
        ["SMB_BUSINESS"] = "Microsoft 365 Apps for business",
        ["OFFICESUBSCRIPTION"] = "Microsoft 365 Apps for enterprise",

        ["ENTERPRISEPACK"] = "Office 365 E3",
        ["DEVELOPERPACK"] = "Office 365 E3 Developer",
        ["ENTERPRISEPREMIUM"] = "Office 365 E5",
        ["STANDARDPACK"] = "Office 365 E1",

        ["SPE_E3"] = "Microsoft 365 E3",
        ["SPE_E5"] = "Microsoft 365 E5",
        ["SPE_F1"] = "Microsoft 365 F3",
        ["SPE_F5_SEC"] = "Microsoft 365 F5 Security + Compliance",

        ["M365EDU_A1"] = "Microsoft 365 A1",
        ["M365EDU_A3_FACULTY"] = "Microsoft 365 A3 for faculty",
        ["M365EDU_A3_STUDENT"] = "Microsoft 365 A3 for students",
        ["M365EDU_A5_FACULTY"] = "Microsoft 365 A5 for faculty",
        ["M365EDU_A5_STUDENT"] = "Microsoft 365 A5 for students",

        ["DESKLESSPACK"] = "Office 365 F3",
        ["DESKLESSWOFFPACK"] = "Office 365 F1",

        ["EMS"] = "Enterprise Mobility + Security E3",
        ["EMSPREMIUM"] = "Enterprise Mobility + Security E5",
        ["AAD_PREMIUM"] = "Microsoft Entra ID P1",
        ["AAD_PREMIUM_P2"] = "Microsoft Entra ID P2",

        ["EXCHANGESTANDARD"] = "Exchange Online Plan 1",
        ["EXCHANGEENTERPRISE"] = "Exchange Online Plan 2",
        ["EXCHANGEARCHIVE"] = "Exchange Online Archiving",
        ["EXCHANGEDESKLESS"] = "Exchange Online Kiosk",

        ["MCOSTANDARD"] = "Skype for Business Online Plan 2",
        ["TEAMS_EXPLORATORY"] = "Microsoft Teams Exploratory",

        ["VISIOCLIENT"] = "Visio Plan 2",
        ["VISIOONLINE_PLAN1"] = "Visio Plan 1",
        ["PROJECTPROFESSIONAL"] = "Project Plan 3",
        ["PROJECTPREMIUM"] = "Project Plan 5",

        ["POWER_BI_STANDARD"] = "Power BI (free)",
        ["POWER_BI_PRO"] = "Power BI Pro",
        ["PBI_PREMIUM_PER_USER"] = "Power BI Premium Per User",

        ["FLOW_FREE"] = "Power Automate Free",
        ["FLOW_PER_USER"] = "Power Automate Premium",
        ["POWERAPPS_VIRAL"] = "Power Apps Developer",
        ["POWERAPPS_PER_USER"] = "Power Apps Premium",

        ["MCOMEETADV"] = "Microsoft 365 Audio Conferencing",
        ["MCOEV"] = "Microsoft Teams Phone Standard",
        ["PHONESYSTEM_VIRTUALUSER"] = "Teams Phone Resource Account",

        ["WINDOWS_STORE"] = "Windows Store for Business",
        ["WIN10_PRO_ENT_SUB"] = "Windows 10/11 Enterprise E3",
        ["WIN10_VDA_E5"] = "Windows 10/11 Enterprise E5",
    };

    public static string Resolve(string skuPartNumber)
    {
        if (string.IsNullOrWhiteSpace(skuPartNumber))
        {
            return string.Empty;
        }

        return FriendlyNames.TryGetValue(skuPartNumber, out var friendly)
            ? friendly
            : HumanizeFallbackName(skuPartNumber);
    }

    private static string HumanizeFallbackName(string skuPartNumber)
    {
        var normalized = skuPartNumber.Trim().Replace('_', ' ');
        normalized = normalized.Replace("O365", "Office 365", StringComparison.OrdinalIgnoreCase);
        normalized = normalized.Replace("M365", "Microsoft 365", StringComparison.OrdinalIgnoreCase);

        return CultureInfo.InvariantCulture.TextInfo.ToTitleCase(normalized.ToLowerInvariant());
    }
}
