using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace OnlyExo365.Shell.Services;

/// <summary>Downloads and converts the Microsoft 365 SKU catalog from Microsoft Learn.</summary>
public sealed partial class LicenseCatalogDownloader : IDisposable
{
    private readonly LicenseCatalogConfiguration _configuration;
    private readonly HttpClient _httpClient;

    // Extracts CSV download link from HTML page
    [GeneratedRegex(
        @"https://download\.microsoft\.com/download/[^""'\s<>]+licensing\.csv",
        RegexOptions.IgnoreCase)]
    private static partial Regex CsvUrlPattern();

    public LicenseCatalogDownloader(LicenseCatalogConfiguration configuration)
    {
        ArgumentNullException.ThrowIfNull(configuration);
        _configuration = configuration;

        _httpClient = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(
                configuration.DownloadTimeoutSeconds > 0
                    ? configuration.DownloadTimeoutSeconds
                    : 30)
        };
        _httpClient.DefaultRequestHeaders.UserAgent.ParseAdd(
            "OnlyExo365/1.0 (LicenseCatalogUpdater)");
    }

    /// <summary>Downloads CSV from documentation page and serializes catalog JSON.</summary>
    public async Task<string> DownloadCatalogJsonAsync(CancellationToken cancellationToken = default)
    {
        var csvUrl = await ResolveCsvUrlAsync(_configuration.RemoteSource, cancellationToken);
        var csvContent = await DownloadCsvAsync(csvUrl, cancellationToken);
        var document = ConvertCsvToDocument(csvContent, _configuration.RemoteSource, csvUrl);
        return JsonSerializer.Serialize(document, LocalSkuCatalogJsonContext.Default.LocalSkuCatalogDocument);
    }

    /// <summary>Resolves CSV URL from documentation HTML.</summary>
    internal async Task<string> ResolveCsvUrlAsync(string documentationUrl, CancellationToken cancellationToken)
    {
        string pageHtml;
        try
        {
            pageHtml = await _httpClient.GetStringAsync(documentationUrl, cancellationToken);
        }
        catch (HttpRequestException ex)
        {
            throw new CatalogDownloadException(
                $"Unable to fetch the Microsoft documentation page '{documentationUrl}': {ex.Message}", ex);
        }

        var match = CsvUrlPattern().Match(pageHtml);
        if (!match.Success)
        {
            throw new CatalogDownloadException(
                $"Could not locate the Microsoft 365 licensing CSV download link on '{documentationUrl}'. " +
                "The page layout may have changed.");
        }

        return match.Value;
    }

    private async Task<string> DownloadCsvAsync(string csvUrl, CancellationToken cancellationToken)
    {
        try
        {
            var bytes = await _httpClient.GetByteArrayAsync(csvUrl, cancellationToken);
            if (bytes.Length == 0)
            {
                throw new CatalogDownloadException(
                    $"The Microsoft 365 licensing CSV downloaded from '{csvUrl}' is empty.");
            }

            // The CSV is UTF-8 with or without BOM.
            return Encoding.UTF8.GetString(bytes).TrimStart('\uFEFF');
        }
        catch (CatalogDownloadException)
        {
            throw;
        }
        catch (HttpRequestException ex)
        {
            throw new CatalogDownloadException(
                $"Failed to download the Microsoft 365 licensing CSV from '{csvUrl}': {ex.Message}", ex);
        }
    }

    /// <summary>Parses CSV and groups rows by SKU into LocalSkuCatalogDocument.</summary>
    internal static LocalSkuCatalogDocument ConvertCsvToDocument(
        string csvContent,
        string sourceUrl,
        string csvDownloadUrl)
    {
        using var reader = new System.IO.StringReader(csvContent);

        // Map column indices from header row
        var headerLine = reader.ReadLine();
        if (string.IsNullOrWhiteSpace(headerLine))
        {
            throw new CatalogDownloadException("The Microsoft 365 licensing CSV has no header row.");
        }

        var headers = ParseCsvRow(headerLine);
        var colGuid = FindColumn(headers, "GUID");
        var colStringId = FindColumn(headers, "String_Id");
        var colProductName = FindColumn(headers, "Product_Display_Name");
        var colServicePlanName = FindColumn(headers, "Service_Plan_Name");
        var colServicePlanId = FindColumn(headers, "Service_Plan_Id");
        var colFriendlyName = FindColumn(headers, "Service_Plans_Included_Friendly_Names");

        // Read all data rows.
        var groups = new Dictionary<string, LocalSkuCatalogEntry>(
            StringComparer.OrdinalIgnoreCase);

        string? line;
        while ((line = reader.ReadLine()) != null)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            var cols = ParseCsvRow(line);

            var skuId = GetColumn(cols, colGuid);
            var skuPartNumber = GetColumn(cols, colStringId);
            var productName = GetColumn(cols, colProductName);
            var servicePlanName = GetColumn(cols, colServicePlanName);
            var servicePlanId = GetColumn(cols, colServicePlanId);
            var friendlyName = GetColumn(cols, colFriendlyName);

            if (string.IsNullOrWhiteSpace(skuId) && string.IsNullOrWhiteSpace(skuPartNumber))
            {
                continue;
            }

            var key = $"{skuPartNumber}|{skuId}";
            if (!groups.TryGetValue(key, out var entry))
            {
                entry = new LocalSkuCatalogEntry
                {
                    SkuId = skuId,
                    SkuPartNumber = skuPartNumber,
                    ProductName = productName
                };
                groups[key] = entry;
            }

            if (string.IsNullOrWhiteSpace(servicePlanName) &&
                string.IsNullOrWhiteSpace(servicePlanId) &&
                string.IsNullOrWhiteSpace(friendlyName))
            {
                continue;
            }

            entry.ServicePlans.Add(new LocalSkuCatalogServicePlan
            {
                ServicePlanName = servicePlanName,
                ServicePlanId = servicePlanId,
                FriendlyName = friendlyName
            });
        }

        if (groups.Count == 0)
        {
            throw new CatalogDownloadException("The Microsoft 365 licensing CSV contains no SKU data rows.");
        }

        var entries = groups.Values
            .Select(e =>
            {
                e.ServicePlans = e.ServicePlans
                    .OrderBy(p => p.ServicePlanName, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(p => p.ServicePlanId, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                return e;
            })
            .OrderBy(e => e.SkuPartNumber, StringComparer.OrdinalIgnoreCase)
            .ThenBy(e => e.SkuId, StringComparer.OrdinalIgnoreCase)
            .ToList();

        return new LocalSkuCatalogDocument
        {
            GeneratedOn = DateTime.UtcNow.ToString("yyyy-MM-dd"),
            Source = sourceUrl,
            CsvDownload = csvDownloadUrl,
            Entries = entries
        };
    }

    /// <summary>Parses CSV row handling quotes and embedded commas.</summary>
    internal static List<string> ParseCsvRow(string line)
    {
        var fields = new List<string>();
        var sb = new StringBuilder();
        var inQuotes = false;

        for (var i = 0; i < line.Length; i++)
        {
            var ch = line[i];

            if (inQuotes)
            {
                if (ch == '"')
                {
                    // Peek ahead for escaped quote ("").
                    if (i + 1 < line.Length && line[i + 1] == '"')
                    {
                        sb.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }
                else
                {
                    sb.Append(ch);
                }
            }
            else
            {
                if (ch == '"')
                {
                    inQuotes = true;
                }
                else if (ch == ',')
                {
                    fields.Add(sb.ToString().Trim());
                    sb.Clear();
                }
                else
                {
                    sb.Append(ch);
                }
            }
        }

        fields.Add(sb.ToString().Trim());
        return fields;
    }

    private static int FindColumn(List<string> headers, string name)
    {
        var idx = headers.FindIndex(h => h.Equals(name, StringComparison.OrdinalIgnoreCase));
        if (idx < 0)
        {
            throw new CatalogDownloadException(
                $"Required column '{name}' not found in the Microsoft 365 licensing CSV header. " +
                $"Available columns: {string.Join(", ", headers)}");
        }

        return idx;
    }

    private static string GetColumn(List<string> cols, int index)
        => index < cols.Count ? cols[index] : string.Empty;

    public void Dispose() => _httpClient.Dispose();
}

public sealed class CatalogDownloadException : Exception
{
    public CatalogDownloadException(string message, Exception? inner = null)
        : base(message, inner)
    {
    }
}

