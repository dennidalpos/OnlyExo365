using System.Text.Json.Serialization;
using OnlyExo365.Contracts.Messages;

namespace OnlyExo365.Contracts.Diagnostics;

public sealed class PersistentLogLevelCounts
{
    [JsonPropertyName("verbose")]
    public int Verbose { get; set; }

    [JsonPropertyName("debug")]
    public int Debug { get; set; }

    [JsonPropertyName("information")]
    public int Information { get; set; }

    [JsonPropertyName("warning")]
    public int Warning { get; set; }

    [JsonPropertyName("error")]
    public int Error { get; set; }

    internal void Increment(LogLevel level)
    {
        switch (level)
        {
            case LogLevel.Verbose:
                Verbose++;
                break;
            case LogLevel.Debug:
                Debug++;
                break;
            case LogLevel.Information:
                Information++;
                break;
            case LogLevel.Warning:
                Warning++;
                break;
            case LogLevel.Error:
                Error++;
                break;
        }
    }
}

