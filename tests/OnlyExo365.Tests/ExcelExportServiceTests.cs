using DocumentFormat.OpenXml.Packaging;
using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;
using OnlyExo365.Shell.Ipc;
using OnlyExo365.Shell.ViewModels;

namespace OnlyExo365.Tests;

public sealed class ExcelExportServiceTests
{
    [Fact]
    public void ResolveExportDirectory_UsesEnvironmentOverride()
    {
        var previousValue = Environment.GetEnvironmentVariable("ONLYEXO365_EXPORT_DIR");
        try
        {
            Environment.SetEnvironmentVariable("ONLYEXO365_EXPORT_DIR", " C:\\exports\\onlyexo365 ");

            var directory = ExcelExportService.ResolveExportDirectory();

            Assert.Equal("C:\\exports\\onlyexo365", directory);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ONLYEXO365_EXPORT_DIR", previousValue);
        }
    }

    [Fact]
    public void ExportWorkbook_WritesHeadersRowsAndNormalizedSheetName()
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"onlyexo365-export-{Guid.NewGuid():N}.xlsx");

        try
        {
            ExcelExportService.ExportWorkbook(
                tempFile,
                "Mailbox/Access:Report*Export[2026]",
                ["User", "Mailbox", "FullAccess"],
                [
                    ["alex@contoso.com", "shared@contoso.com", "Si"],
                    ["bianca@contoso.com", "ops@contoso.com", "-"]
                ]);

            using var spreadsheet = SpreadsheetDocument.Open(tempFile, false);
            var workbook = spreadsheet.WorkbookPart?.Workbook;
            var sheet = Assert.Single(workbook?.Sheets?.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>() ?? []);
            Assert.Equal("Mailbox-Access-Report-Export-20", sheet.Name?.Value);

            var worksheetPart = (WorksheetPart)spreadsheet.WorkbookPart!.GetPartById(sheet.Id!);
            var rows = worksheetPart.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>().ToList();

            Assert.Equal(3, rows.Count);
            Assert.Equal("User", GetInlineText(rows[0].Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().ElementAt(0)));
            Assert.Equal("Mailbox", GetInlineText(rows[0].Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().ElementAt(1)));
            Assert.Equal("alex@contoso.com", GetInlineText(rows[1].Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().ElementAt(0)));
            Assert.Equal("shared@contoso.com", GetInlineText(rows[1].Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().ElementAt(1)));
            Assert.Equal("-", GetInlineText(rows[2].Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().ElementAt(2)));
        }
        finally
        {
            if (File.Exists(tempFile))
            {
                File.Delete(tempFile);
            }
        }
    }

    [Fact]
    public void MailboxSpaceExportCommand_CanExecuteOnlyWhenRowsAreAvailable()
    {
        using var shell = CreateConnectedShell();
        var viewModel = new MailboxSpaceViewModel(new MailboxReportWorkerService(), new NavigationService(), shell);

        Assert.False(viewModel.ExportExcelCommand.CanExecute(null));

        viewModel.Mailboxes.Add(new MailboxSpaceItemViewModel(new MailboxSpaceItemDto
        {
            Identity = "mailbox-01",
            DisplayName = "Mailbox 01",
            PrimarySmtpAddress = "mailbox01@contoso.com",
            TotalItemSize = "10 GB",
            ProhibitSendQuota = "50 GB",
            ProhibitSendQuotaBytes = 50L * 1024 * 1024 * 1024,
            TotalItemSizeBytes = 10L * 1024 * 1024 * 1024
        }));

        Assert.True(viewModel.ExportExcelCommand.CanExecute(null));
    }

    [Fact]
    public void MailboxAccessExportCommand_CanExecuteOnlyWhenRowsAreAvailable()
    {
        using var shell = CreateConnectedShell();
        var viewModel = new MailboxAccessReportViewModel(new MailboxReportWorkerService(), shell);

        Assert.False(viewModel.ExportExcelCommand.CanExecute(null));

        viewModel.Rows.Add(new MailboxAccessMatrixRowViewModel(
            "alex@contoso.com",
            "shared@contoso.com",
            [
                new MailboxAccessGrantDto
                {
                    User = "alex@contoso.com",
                    MailboxPrimarySmtpAddress = "shared@contoso.com",
                    MailboxIdentity = "shared",
                    PermissionType = "FullAccess",
                    AccessRights = ["FullAccess"]
                }
            ]));

        Assert.True(viewModel.ExportExcelCommand.CanExecute(null));
    }

    private static ShellViewModel CreateConnectedShell()
    {
        var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        typeof(ShellViewModel).GetProperty(nameof(ShellViewModel.ExchangeState))!.SetValue(shell, ConnectionState.Connected);
        typeof(ShellViewModel).GetProperty(nameof(ShellViewModel.Capabilities))!.SetValue(shell, new CapabilityMapDto());
        return shell;
    }

    private static string GetInlineText(DocumentFormat.OpenXml.Spreadsheet.Cell cell)
        => cell.InlineString?.Text?.Text ?? string.Empty;

    private sealed class MailboxReportWorkerService : TestMailboxesWorkerServiceBase
    {
    }
}

