using System.Globalization;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExchangeAdmin.Presentation.Services;

internal static class ExcelExportService
{
    public static string ResolveExportDirectory()
    {
        var fromEnv = Environment.GetEnvironmentVariable("EXCHANGEADMIN_EXPORT_DIR");
        if (!string.IsNullOrWhiteSpace(fromEnv))
        {
            return fromEnv.Trim();
        }

        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OnlyExo365",
            "exports");
    }

    public static void ExportWorkbook(
        string filePath,
        string sheetName,
        IReadOnlyList<string> headers,
        IEnumerable<IReadOnlyList<string?>> rows)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        ArgumentNullException.ThrowIfNull(headers);
        ArgumentNullException.ThrowIfNull(rows);

        using var spreadsheet = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(new SheetViews(new SheetView { WorkbookViewId = 0U }), sheetData);

        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = CreateStylesheet();
        stylesPart.Stylesheet.Save();

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = NormalizeSheetName(sheetName)
        });

        var headerRow = new Row();
        foreach (var title in headers)
        {
            headerRow.Append(CreateTextCell(title, 1U));
        }

        sheetData.Append(headerRow);

        foreach (var rowValues in rows)
        {
            var row = new Row();
            foreach (var value in rowValues)
            {
                row.Append(CreateTextCell(value));
            }

            sheetData.Append(row);
        }

        worksheetPart.Worksheet.Save();
        workbookPart.Workbook.Save();
    }

    private static string NormalizeSheetName(string? sheetName)
    {
        const int maxLength = 31;
        var normalized = string.IsNullOrWhiteSpace(sheetName)
            ? "Export"
            : sheetName.Trim();

        foreach (var invalidCharacter in new[] { '\\', '/', '?', '*', '[', ']', ':' })
        {
            normalized = normalized.Replace(invalidCharacter, '-');
        }

        return normalized.Length <= maxLength
            ? normalized
            : normalized[..maxLength];
    }

    private static Stylesheet CreateStylesheet()
    {
        var fonts = new Fonts(
            new Font(),
            new Font(new Bold()));

        var fills = new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),
            new Fill(new PatternFill(new ForegroundColor { Rgb = HexBinaryValue.FromString("FFDDEBF7") }) { PatternType = PatternValues.Solid }));

        var borders = new Borders(new Border());

        var cellFormats = new CellFormats(
            new CellFormat(),
            new CellFormat
            {
                FontId = 1,
                FillId = 2,
                BorderId = 0,
                ApplyFont = true,
                ApplyFill = true
            });

        return new Stylesheet(fonts, fills, borders, cellFormats);
    }

    private static Cell CreateTextCell(string? value, uint styleIndex = 0)
    {
        return new Cell
        {
            DataType = CellValues.InlineString,
            StyleIndex = styleIndex,
            InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text(value ?? string.Empty))
        };
    }
}
