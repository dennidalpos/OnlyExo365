using System.Text;
using OnlyExo365.Contracts;

namespace OnlyExo365.Tests;

public class IpcLineReaderTests
{
    [Fact]
    public async Task ReadLineAsync_HandlesBomCrLfAndBufferedNextLine()
    {
        var bytes = Encoding.UTF8.GetPreamble()
            .Concat(Encoding.UTF8.GetBytes("{\"name\":\"é\"}\r\n{\"ok\":true}\n"))
            .ToArray();
        using var stream = new MemoryStream(bytes);
        var reader = new IpcLineReader(stream);

        var first = await reader.ReadLineAsync();
        var second = await reader.ReadLineAsync();

        Assert.Equal("{\"name\":\"é\"}", first?.Line);
        Assert.Equal(Encoding.UTF8.GetByteCount(first!.Value.Line!), first.Value.SizeBytes);
        Assert.Equal("{\"ok\":true}", second?.Line);
        Assert.Null(await reader.ReadLineAsync());
    }

    [Fact]
    public async Task ReadLineAsync_AcceptsExactUtf8Limit()
    {
        var message = new string('é', IpcConstants.MaxMessageSizeBytes / 2);
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(message + "\n"));
        var reader = new IpcLineReader(stream);

        var result = await reader.ReadLineAsync();

        Assert.False(result?.IsTooLarge);
        Assert.Equal(IpcConstants.MaxMessageSizeBytes, result?.SizeBytes);
        Assert.Equal(message, result?.Line);
    }

    [Fact]
    public async Task ReadLineAsync_DiscardsOversizeMultibyteLineAndReadsNextLine()
    {
        var oversized = new string('é', IpcConstants.MaxMessageSizeBytes / 2 + 1);
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(oversized + "\n{\"ok\":true}\n"));
        var reader = new IpcLineReader(stream);

        var rejected = await reader.ReadLineAsync();
        var next = await reader.ReadLineAsync();

        Assert.True(rejected?.IsTooLarge);
        Assert.Null(rejected?.Line);
        Assert.Equal(IpcConstants.MaxMessageSizeBytes + 2, rejected?.SizeBytes);
        Assert.Equal("{\"ok\":true}", next?.Line);
    }
}
