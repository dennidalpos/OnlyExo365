using OnlyExo365.Contracts;

namespace OnlyExo365.Tests;

public class AtomicFileWriterTests
{
    [Fact]
    public void Write_PreservesExistingFileAndRemovesPartialOutputOnFailure()
    {
        var directory = Path.Combine(Path.GetTempPath(), "OnlyExo365.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var destination = Path.Combine(directory, "report.xlsx");

        try
        {
            File.WriteAllText(destination, "previous report");

            Assert.Throws<InvalidOperationException>(() => AtomicFileWriter.Write(destination, temporaryPath =>
            {
                File.WriteAllText(temporaryPath, "partial report");
                throw new InvalidOperationException("generation failed");
            }));

            Assert.Equal("previous report", File.ReadAllText(destination));
            Assert.Equal([destination], Directory.GetFiles(directory));

            AtomicFileWriter.Write(destination, temporaryPath => File.WriteAllText(temporaryPath, "new report"));
            Assert.Equal("new report", File.ReadAllText(destination));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }
}
