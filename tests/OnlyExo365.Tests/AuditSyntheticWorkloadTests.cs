using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Management.Automation;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OnlyExo365.Contracts.Diagnostics;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Shell.Helpers;
using OnlyExo365.Shell.Services;
using OnlyExo365.Worker.PowerShell;

namespace OnlyExo365.Tests;

public sealed class AuditSyntheticWorkloadTests(ITestOutputHelper output)
{
    [Fact]
    public void PopulatedData_MappingCollectionsLogsAndExportPreserveAllRows()
    {
        var root = Path.Combine(Path.GetTempPath(), "OnlyExo365.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            var measurements = new Dictionary<string, double>();
            var records = Enumerable.Range(0, 10000).Select(i =>
            {
                var record = new PSObject();
                record.Properties.Add(new PSNoteProperty("Identity", $"record-{i}"));
                record.Properties.Add(new PSNoteProperty("Operations", "MailItemsAccessed"));
                record.Properties.Add(new PSNoteProperty("AuditData", "{\"synthetic\":true}"));
                return record;
            }).ToArray();
            var watch = Stopwatch.StartNew();
            var mapped = ExoComplianceCommands.MapSearchUnifiedAuditLogResponse(records, records.Length);
            measurements["map_10000_ms"] = watch.Elapsed.TotalMilliseconds;
            Assert.Equal(records.Length, mapped.Results.Count);
            Assert.Equal("record-9999", mapped.Results[^1].Identity);

            var collection = new ObservableCollection<UnifiedAuditLogRecordDto>();
            watch.Restart();
            collection.ReplaceAll(mapped.Results);
            collection.ReplaceAll(mapped.Results.Take(5000));
            measurements["collections_10000_then_5000_ms"] = watch.Elapsed.TotalMilliseconds;
            Assert.Equal(5000, collection.Count);
            Assert.Equal("record-4999", collection[^1].Identity);

            var now = DateTime.UtcNow;
            var logs = new PersistentLogWriter("synthetic", Path.Combine(root, "logs"), () => now);
            watch.Restart();
            for (var i = 0; i < 1000; i++) logs.Write(LogLevel.Information, "probe", $"Synthetic entry {i}");
            var summary = PersistentLogStore.BuildSummary(logs.LogDirectoryPath, now);
            measurements["log_write_and_summary_1000_ms"] = watch.Elapsed.TotalMilliseconds;
            Assert.Equal(1000, summary.TotalEntries);
            Assert.Equal(0, summary.ParseErrors);

            var path = Path.Combine(root, "synthetic.xlsx");
            watch.Restart();
            ExcelExportService.ExportWorkbook(path, "Synthetic", ["Identity", "Literal"],
                mapped.Results.Select(row => (IReadOnlyList<string?>)new[] { row.Identity, "=1+1" }));
            measurements["export_10000_ms"] = watch.Elapsed.TotalMilliseconds;
            using var workbook = SpreadsheetDocument.Open(path, false);
            var sheet = Assert.Single(workbook.WorkbookPart!.WorksheetParts).Worksheet!;
            var rows = sheet.Descendants<Row>().ToArray();
            Assert.Equal(10001, rows.Length);
            var literal = rows[1].Elements<Cell>().Last();
            Assert.Null(literal.CellFormula);
            Assert.Equal("=1+1", literal.InlineString!.InnerText);
            output.WriteLine(JsonSerializer.Serialize(measurements));
        }
        finally { Directory.Delete(root, recursive: true); }
    }
}
