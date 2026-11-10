using System.IO;
using System.Text;

namespace OnlyExo365.Contracts;

public readonly record struct IpcLineReadResult(string? Line, long SizeBytes, bool IsTooLarge);

public sealed class IpcLineReader
{
    private readonly Stream _stream;
    private readonly byte[] _buffer = new byte[IpcConstants.MaxReadBufferSize];
    private int _position;
    private int _count;
    private bool _firstLine = true;

    public IpcLineReader(Stream stream)
    {
        _stream = stream ?? throw new ArgumentNullException(nameof(stream));
    }

    public async ValueTask<IpcLineReadResult?> ReadLineAsync(CancellationToken cancellationToken = default)
    {
        using var line = new MemoryStream();
        var sizeBytes = 0L;
        var tooLarge = false;
        var maxBufferedBytes = IpcConstants.MaxMessageSizeBytes + (_firstLine ? 3 : 0) + 1;

        while (true)
        {
            if (_position == _count)
            {
                _count = await _stream.ReadAsync(_buffer, cancellationToken).ConfigureAwait(false);
                _position = 0;
                if (_count == 0)
                {
                    return sizeBytes == 0 ? null : FinishLine(line, sizeBytes, tooLarge);
                }
            }

            var remaining = _buffer.AsSpan(_position, _count - _position);
            var newline = remaining.IndexOf((byte)'\n');
            var segmentLength = newline < 0 ? remaining.Length : newline;

            sizeBytes += segmentLength;
            if (!tooLarge)
            {
                if (line.Length + segmentLength > maxBufferedBytes)
                {
                    tooLarge = true;
                }
                else
                {
                    line.Write(remaining[..segmentLength]);
                }
            }

            _position += segmentLength + (newline >= 0 ? 1 : 0);
            if (newline >= 0)
            {
                return FinishLine(line, sizeBytes, tooLarge);
            }
        }
    }

    private IpcLineReadResult FinishLine(MemoryStream line, long sizeBytes, bool tooLarge)
    {
        var bytes = line.GetBuffer().AsSpan(0, (int)line.Length);
        var preamble = Encoding.UTF8.GetPreamble();
        var hasBom = _firstLine && bytes.StartsWith(preamble);
        _firstLine = false;

        if (hasBom)
        {
            bytes = bytes[preamble.Length..];
            sizeBytes -= preamble.Length;
        }

        if (!tooLarge && bytes.EndsWith("\r"u8))
        {
            bytes = bytes[..^1];
            sizeBytes--;
        }

        if (tooLarge || !IpcConstants.IsValidMessageSize(sizeBytes))
        {
            return new IpcLineReadResult(null, sizeBytes, true);
        }

        return new IpcLineReadResult(Encoding.UTF8.GetString(bytes), sizeBytes, false);
    }
}
