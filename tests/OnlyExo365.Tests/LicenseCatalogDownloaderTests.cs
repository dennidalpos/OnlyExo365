using OnlyExo365.Shell.Services;

namespace OnlyExo365.Tests;

public sealed class LicenseCatalogDownloaderTests
{
    // -------------------------------------------------------------------------
    // ParseCsvRow
    // -------------------------------------------------------------------------

    [Fact]
    public void ParseCsvRow_SplitsSimpleRow()
    {
        var result = LicenseCatalogDownloader.ParseCsvRow("a,b,c");
        Assert.Equal(["a", "b", "c"], result);
    }

    [Fact]
    public void ParseCsvRow_HandlesQuotedField()
    {
        var result = LicenseCatalogDownloader.ParseCsvRow("\"hello, world\",b,c");
        Assert.Equal(3, result.Count);
        Assert.Equal("hello, world", result[0]);
    }

    [Fact]
    public void ParseCsvRow_HandlesEscapedQuoteInsideQuotedField()
    {
        var result = LicenseCatalogDownloader.ParseCsvRow("\"say \"\"hi\"\"\",b");
        Assert.Equal(2, result.Count);
        Assert.Equal("say \"hi\"", result[0]);
    }

    [Fact]
    public void ParseCsvRow_TrimsUnquotedFields()
    {
        var result = LicenseCatalogDownloader.ParseCsvRow(" a , b , c ");
        Assert.Equal(["a", "b", "c"], result);
    }

    // -------------------------------------------------------------------------
    // ConvertCsvToDocument
    // -------------------------------------------------------------------------

    private static string BuildMinimalCsv(params (string Guid, string StringId, string ProductName)[] rows)
    {
        var sb = new System.Text.StringBuilder();
        sb.AppendLine("Product_Display_Name,String_Id,GUID,Service_Plan_Name,Service_Plan_Id,Service_Plans_Included_Friendly_Names");
        foreach (var (guid, stringId, name) in rows)
        {
            sb.AppendLine($"{name},{stringId},{guid},SOME_PLAN,plan-id-1,Some Plan");
        }

        return sb.ToString();
    }

    private static string BuildCsvWithServicePlans(
        params (string Guid, string StringId, string ProductName, string ServicePlanName, string ServicePlanId, string FriendlyName)[] rows)
    {
        var sb = new System.Text.StringBuilder();
        sb.AppendLine("Product_Display_Name,String_Id,GUID,Service_Plan_Name,Service_Plan_Id,Service_Plans_Included_Friendly_Names");
        foreach (var row in rows)
        {
            sb.AppendLine(
                $"{row.ProductName},{row.StringId},{row.Guid},{row.ServicePlanName},{row.ServicePlanId},{row.FriendlyName}");
        }

        return sb.ToString();
    }

    [Fact]
    public void ConvertCsvToDocument_ProducesCorrectEntries()
    {
        var csv = BuildMinimalCsv(
            ("guid-a", "SKU_A", "Product A"),
            ("guid-b", "SKU_B", "Product B"));

        var doc = LicenseCatalogDownloader.ConvertCsvToDocument(
            csv, "https://source.example", "https://csv.example");

        Assert.Equal(2, doc.Entries.Count);

        var entryA = doc.Entries.First(e => e.SkuPartNumber == "SKU_A");
        Assert.Equal("guid-a", entryA.SkuId);
        Assert.Equal("Product A", entryA.ProductName);
    }

    [Fact]
    public void ConvertCsvToDocument_GroupsServicePlansWithSameSkuId()
    {
        var csv = BuildCsvWithServicePlans(
            ("guid-a", "SKU_A", "Product A", "PLAN_B", "plan-id-b", "Plan B"),
            ("guid-a", "SKU_A", "Product A", "PLAN_A", "plan-id-a", "Plan A"));

        var doc = LicenseCatalogDownloader.ConvertCsvToDocument(
            csv, "https://source.example", "https://csv.example");

        var entry = Assert.Single(doc.Entries);
        Assert.Equal("SKU_A", entry.SkuPartNumber);
        Assert.Collection(
            entry.ServicePlans,
            servicePlan =>
            {
                Assert.Equal("PLAN_A", servicePlan.ServicePlanName);
                Assert.Equal("plan-id-a", servicePlan.ServicePlanId);
                Assert.Equal("Plan A", servicePlan.FriendlyName);
            },
            servicePlan =>
            {
                Assert.Equal("PLAN_B", servicePlan.ServicePlanName);
                Assert.Equal("plan-id-b", servicePlan.ServicePlanId);
                Assert.Equal("Plan B", servicePlan.FriendlyName);
            });
    }

    [Fact]
    public void ConvertCsvToDocument_SkipsBlankServicePlanRows()
    {
        var csv = BuildCsvWithServicePlans(
            ("guid-a", "SKU_A", "Product A", string.Empty, string.Empty, string.Empty));

        var doc = LicenseCatalogDownloader.ConvertCsvToDocument(
            csv, "https://source.example", "https://csv.example");

        var entry = Assert.Single(doc.Entries);
        Assert.Empty(entry.ServicePlans);
    }

    [Fact]
    public void ConvertCsvToDocument_SetsGeneratedOnToToday()
    {
        var csv = BuildMinimalCsv(("guid-a", "SKU_A", "Product A"));

        var doc = LicenseCatalogDownloader.ConvertCsvToDocument(
            csv, "https://source.example", "https://csv.example");

        Assert.Equal(DateTime.UtcNow.ToString("yyyy-MM-dd"), doc.GeneratedOn);
    }

    [Fact]
    public void ConvertCsvToDocument_ThrowsCatalogDownloadException_WhenHeaderMissingColumn()
    {
        // CSV with no GUID column.
        const string csv = "Product_Display_Name,String_Id,Service_Plan_Name\nProduct A,SKU_A,PLAN\n";

        Assert.Throws<CatalogDownloadException>(
            () => LicenseCatalogDownloader.ConvertCsvToDocument(
                csv, "https://source.example", "https://csv.example"));
    }

    [Fact]
    public void ConvertCsvToDocument_ThrowsCatalogDownloadException_WhenCsvEmpty()
    {
        Assert.Throws<CatalogDownloadException>(
            () => LicenseCatalogDownloader.ConvertCsvToDocument(
                string.Empty, "https://source.example", "https://csv.example"));
    }

    [Fact]
    public void ConvertCsvToDocument_ThrowsCatalogDownloadException_WhenNoDataRows()
    {
        // Only header, no data.
        const string csv = "Product_Display_Name,String_Id,GUID,Service_Plan_Name,Service_Plan_Id,Service_Plans_Included_Friendly_Names\n";

        Assert.Throws<CatalogDownloadException>(
            () => LicenseCatalogDownloader.ConvertCsvToDocument(
                csv, "https://source.example", "https://csv.example"));
    }

    [Fact]
    public void ConvertCsvToDocument_SetsSourceAndCsvDownloadFields()
    {
        var csv = BuildMinimalCsv(("guid-a", "SKU_A", "Product A"));

        var doc = LicenseCatalogDownloader.ConvertCsvToDocument(
            csv, "https://my.source", "https://my.csv");

        Assert.Equal("https://my.source", doc.Source);
        Assert.Equal("https://my.csv", doc.CsvDownload);
    }
}

