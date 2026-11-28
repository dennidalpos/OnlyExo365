using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Worker.PowerShell;

namespace OnlyExo365.Tests;

public sealed class AuditMessageTraceRegressionTests
{
    [Theory]
    [InlineData(0, 1, 10, 0, false, true, 1)]
    [InlineData(5, 1, 10, 5, false, true, 1)]
    [InlineData(10000, 1, 10, 10, true, false, 1)]
    [InlineData(5005, 1001, 5, 5, false, true, 2)]
    public async Task TraceQuery_LoadsOnlyRequestedWindow(int total, int page, int pageSize, int expectedRows, bool hasMore, bool exact, int calls)
    {
        using var fixture = new LocalPowerShellFixture();
        await fixture.DefineCommandsAsync($$"""
            $global:traceOffset = 0
            $global:traceCalls = 0
            $global:traceNow = [datetime]::UtcNow
            function Get-MessageTraceV2 {
                param([int]$ResultSize)
                $global:traceCalls++
                $count = [Math]::Min($ResultSize, {{total}} - $global:traceOffset)
                for ($index = 0; $index -lt $count; $index++) {
                    $global:traceOffset++
                    [pscustomobject]@{
                        MessageId = [string]$global:traceOffset
                        MessageTraceId = '00000000-0000-0000-0000-000000000001'
                        Received = $global:traceNow.AddSeconds(-$global:traceOffset)
                        RecipientAddress = 'dummy@example.com'
                    }
                }
            }
            """);
        var response = await new ExoMessageTraceCommands(fixture.Engine).GetMessageTraceAsync(new GetMessageTraceRequest
        {
            StartDate = DateTime.UtcNow.AddDays(-7), EndDate = DateTime.UtcNow.AddDays(-1), Page = page, PageSize = pageSize
        });
        Assert.Equal(expectedRows, response.Messages.Count);
        Assert.Equal(hasMore, response.HasMore);
        Assert.Equal(exact, response.IsTotalCountExact);
        var callCount = await fixture.Engine.ExecuteAsync("$global:traceCalls");
        Assert.Equal(calls, Convert.ToInt32(callCount.Output.Single().BaseObject));
    }

    [Theory]
    [InlineData(-12, -1, 1, 100)]
    [InlineData(-92, -91, 1, 100)]
    [InlineData(-1, -2, 1, 100)]
    [InlineData(-2, -1, 0, 100)]
    [InlineData(-2, -1, 1, 0)]
    public async Task InvalidTraceRequest_IsRejectedBeforeExecution(int startDays, int endDays, int page, int pageSize)
    {
        using var fixture = new LocalPowerShellFixture();
        await Assert.ThrowsAnyAsync<ArgumentException>(async () => await new ExoMessageTraceCommands(fixture.Engine).GetMessageTraceAsync(new GetMessageTraceRequest
        {
            StartDate = DateTime.UtcNow.AddDays(startDays), EndDate = DateTime.UtcNow.AddDays(endDays), Page = page, PageSize = pageSize
        }));
    }

    [Fact]
    public async Task StalledCursor_FailsInsteadOfLooping()
    {
        using var fixture = new LocalPowerShellFixture();
        await fixture.DefineCommandsAsync("""
            $global:traceNow = [datetime]::UtcNow.AddDays(-1)
            function Get-MessageTraceV2 {
                param([int]$ResultSize)
                for ($index = 0; $index -lt $ResultSize; $index++) {
                    [pscustomobject]@{ Received = $global:traceNow; RecipientAddress = 'dummy@example.com' }
                }
            }
            """);
        var error = await Assert.ThrowsAsync<InvalidOperationException>(async () => await new ExoMessageTraceCommands(fixture.Engine).GetMessageTraceAsync(new GetMessageTraceRequest
        {
            StartDate = DateTime.UtcNow.AddDays(-7), EndDate = DateTime.UtcNow, Page = 3, PageSize = 5000
        }));
        Assert.Contains("cursor did not advance", error.Message, StringComparison.Ordinal);
    }
}
