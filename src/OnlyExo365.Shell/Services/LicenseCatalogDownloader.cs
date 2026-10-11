using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic.FileIO;

namespace OnlyExo365.Shell.Services;

/// <summary>Downloads Microsoft Learn SKU mappings.</summary>
public sealed partial class LicenseCatalogDownloader : IDisposable
{
    private readonly LicenseCatalogConfiguration _configuration;
    private readonly HttpClient _httpClient;

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

    /// <summary>Downloads CSV and returns catalog JSON.</summary>
    public async Task<string> DownloadCatalogJsonAsync(CancellationToken cancellationToken = default)
    {
        var csvUrl = await ResolveCsvUrlAsync(_configuration.RemoteSource, cancellationToken);
        var csvContent = await DownloadCsvAsync(csvUrl, cancellationToken);
        var document = ConvertCsvToDocument(csvContent, _configuration.RemoteSource, csvUrl);
        return JsonSerializer.Serialize(document, LocalSkuCatalogJsonContext.Default.LocalSkuCatalogDocument);
    }

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

            // Accept UTF-8 with/without BOM.
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

    /// <summary>Groups CSV rows by SKU.</summary>
    internal static LocalSkuCatalogDocument ConvertCsvToDocument(
        string csvContent,
        string sourceUrl,
        string csvDownloadUrl)
    {
        using var reader = new TextFieldParser(new System.IO.StringReader(csvContent))
        {
            TextFieldType = FieldType.Delimited,
            HasFieldsEnclosedInQuotes = true,
            TrimWhiteSpace = true
        };
        reader.SetDelimiters(",");

        var headers = ReadCsvFields(reader);
        if (headers is null || headers.Count == 0 || headers.All(string.IsNullOrWhiteSpace))
        {
            throw new CatalogDownloadException("The Microsoft 365 licensing CSV has no header row.");
        }

        headers[0] = headers[0].TrimStart('\uFEFF');
        var colGuid = FindColumn(headers, "GUID");
        var colStringId = FindColumn(headers, "String_Id");
        var colProductName = FindColumn(headers, "Product_Display_Name");
        var colServicePlanName = FindColumn(headers, "Service_Plan_Name");
        var colServicePlanId = FindColumn(headers, "Service_Plan_Id");
        var colFriendlyName = FindColumn(headers, "Service_Plans_Included_Friendly_Names");

        var groups = new Dictionary<string, LocalSkuCatalogEntry>(
            StringComparer.OrdinalIgnoreCase);

        while (!reader.EndOfData)
        {
            var cols = ReadCsvFields(reader);
            if (cols is null)
            {
                break;
            }

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

    private static List<string>? ReadCsvFields(TextFieldParser reader)
    {
        try
        {
            return reader.ReadFields()?.ToList();
        }
        catch (MalformedLineException ex)
        {
            throw new CatalogDownloadException(
                $"The Microsoft 365 licensing CSV contains a malformed row at line {reader.ErrorLineNumber}.", ex);
        }
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

